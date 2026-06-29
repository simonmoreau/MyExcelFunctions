using Azure.AI.OpenAI;
using ExcelDna.Integration;
using ExcelFunctions.AI;
using ExcelFunctions.Services;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using OllamaSharp;
using OpenAI;
using Serilog;
using Serilog.Extensions.Logging;
using System.ClientModel;
using System.Net;
using System.Windows.Input;

namespace ExcelFunctions
{
    internal static class ContainerOperations
    {
        //Excel needs some extras help in only registering dependencies once
        private static readonly Lazy<IServiceProvider> ContainerSingleton = new(() => CreateContainer());
        public static IServiceProvider Container => ContainerSingleton.Value;

        //The DI registrations
        internal static IServiceProvider CreateContainer(string? basePath = null)
        {
            ServiceCollection container = new ServiceCollection();

            basePath ??= ExcelDnaUtil.XllPathInfo?.Directory?.FullName ??
                throw new Exception($"Unable to configure app, invalid value for ExcelDnaUtil.XllPathInfo='{ExcelDnaUtil.XllPathInfo}'");

            IConfiguration configuration = new ConfigurationBuilder()
                .SetBasePath(basePath)
                .AddJsonFile("appsettings.json")
#if DEBUG
                .AddJsonFile("appsettings.Development.json", true)
#endif
                .Build();

            //Settings

            var settings = configuration.GetSection(nameof(AppSettings)).Get<AppSettings>();

            if (settings == null) throw new Exception("No configuration section found called AppSettings");
            container.AddSingleton(_ => settings);

            //Logging
            container.AddSingleton(_ => ConfigureLogging(configuration));
            container.AddSingleton(sp => sp.GetRequiredService<ILoggerFactory>().CreateLogger("ExcelFunctions"));

            // AI
            container.AddAIServices(configuration);

            // HttpClient
            container.AddHttpClient<ExcelHttpClient>(client =>
            {
                client.DefaultRequestHeaders.Accept.Clear();
                client.Timeout = Timeout.InfiniteTimeSpan;
                client.DefaultRequestHeaders.Add("User-Agent", "ExcelHttpClient/0.0");
            })
                .ConfigurePrimaryHttpMessageHandler(() => new HttpClientHandler()
            {
                AutomaticDecompression = DecompressionMethods.GZip,
                MaxConnectionsPerServer = 200
            });

            return container.BuildServiceProvider();
        }

        private static ILoggerFactory ConfigureLogging(IConfiguration configuration)
        {
            IConfigurationSection config = configuration.GetSection("AppSettings");
            string appVersion = config["Version"] ?? "Unknown Version";
            var serilog = new LoggerConfiguration()
                .ReadFrom.Configuration(config)
                .Enrich.WithProperty("AppName", "ExcelFunctions")
                .Enrich.WithProperty("AppVersion", appVersion)
                .CreateLogger();

            return new LoggerFactory(new[] { new SerilogLoggerProvider(serilog) });
        }

        public static IServiceCollection AddAIServices(this IServiceCollection services, IConfiguration configuration)
        {
            services.AddSingleton<RateLimitingHandler>();

            OllamaSettings? ollamaSettings = configuration.GetSection(nameof(OllamaSettings)).Get<OllamaSettings>();

            //if (ollamaSettings != null)
            //{
            //    Uri ollamaEndpoint = new Uri(ollamaSettings.Endpoint);
            //    string modelId = ollamaSettings.ModelId;
            //    IChatClient ollama = new OllamaApiClient(ollamaEndpoint)
            //    {
            //        SelectedModel = modelId
            //    };
            //    services.AddChatClient(ollama);
            //    return services;
            //}

            OpenAISettings? openAISettings = configuration.GetSection(nameof(OpenAISettings)).Get<OpenAISettings>();

            if (openAISettings != null)
            {
                string modelId = openAISettings.ModelId;
                IChatClient client = new OpenAIClient(openAISettings.Key).GetChatClient(modelId).AsIChatClient();
                services.AddChatClient(client);
                return services;
            }

            AzureOpenAI? azureOpenAISettigns = configuration.GetSection(nameof(AzureOpenAI)).Get<AzureOpenAI>();

            string deploymentName = azureOpenAISettigns.ModelId;
            Uri endpoint = new Uri(azureOpenAISettigns.Endpoint);
            ApiKeyCredential apiKey = new ApiKeyCredential(azureOpenAISettigns.Key);

            // Build AI client and chat client with defensive checks
            AzureOpenAIClient azureClient = new(endpoint, apiKey);
            if (string.IsNullOrWhiteSpace(deploymentName))
                throw new InvalidOperationException("Azure OpenAI deployment name (ModelId) is not configured.");
            IChatClient chatClient;
            try
            {
                chatClient = azureClient.GetChatClient(deploymentName).AsIChatClient();
            }
            catch (Exception ex)
            {
                // Surface a clear error in the logs to aid debugging when the add-in starts
                Serilog.Log.Logger?.Error(ex, "Failed to create IChatClient for deployment '{Deployment}'", deploymentName);
                throw;
            }

            services.AddChatClient(chatClient);

            return services;
        }
    }
}
