using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Serilog;
using ColetaAutomatizadaTG.Interfaces;
using ColetaAutomatizadaTG.Data;
using ColetaAutomatizadaTG.Handler;
using ColetaAutomatizadaTG.Services;

namespace ColetaAutomatizadaTG
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            var builder = Host.CreateDefaultBuilder(args)
                .UseWindowsService()
                .ConfigureAppConfiguration((context, config) =>
                {
                    config.AddJsonFile(
                        "appsettings.json",
                        optional: false,
                        reloadOnChange: true
                    );
                })
                .UseSerilog((context, services, loggerConfig) =>
                {
                    loggerConfig
                        .ReadFrom.Configuration(context.Configuration) 
                        .ReadFrom.Services(services)                  
                        .Enrich.FromLogContext();
                })
                .ConfigureServices((context, services) =>
                {
                    IConfiguration configuration = context.Configuration;

                    services.AddSingleton<ExcelExporter>();
                    services.AddSingleton<IAutomacaoRepository, AutomacaoRepository>();
                    services.AddSingleton<ISeleniumHandler, SeleniumHandler>();
                    services.AddHostedService<Worker>();

                });

            var host = builder.Build();
            await host.RunAsync();
        }
    }
}
