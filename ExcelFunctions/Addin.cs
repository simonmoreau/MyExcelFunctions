using System;
using System.Net;
using System.Threading;
using ExcelDna.Integration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;

namespace ExcelFunctions
{
    public class AddIn : IExcelAddIn
    {
        // ReSharper disable once MemberCanBePrivate.Global
        public static SynchronizationContext? SynchronizationContext { get; private set; }

        public void AutoOpen()
        {
            SynchronizationContext = new ExcelSynchronizationContext();
            SynchronizationContext.SetSynchronizationContext(SynchronizationContext);
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            try
            {
                IServiceProvider container = ContainerOperations.Container; //all dependency registrations happen here

                ILogger logger = container.GetService<ILogger>() ?? throw new Exception("Could not resolve an ILogger");

                logger.LogInformation("Loading AzureServiceBus.ExcelAddIn");

                ExcelIntegration.RegisterUnhandledExceptionHandler(ex =>
                {
                    // ReSharper disable once InvertIf
                    if (ex is Exception exception)
                    {
                        logger.LogError(exception, "Unhandled Exception");
                    }
                    return ex;
                });

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Critical error in AutoOpen: {ex}");
                throw;
            }
        }

        public void AutoClose()
        {
            try
            {
                ILogger? logger = ContainerOperations.Container.GetService<ILogger>();
                logger?.LogInformation("Unloading AzureServiceBus.ExcelAddIn");
            }
            catch
            {
                // ignore
            }
        }
    }
}

