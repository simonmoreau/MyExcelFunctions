using ExcelDna.Integration;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Serilog;
using Serilog.Extensions.Logging;

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
    }
}
