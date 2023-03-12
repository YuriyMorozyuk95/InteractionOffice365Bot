using System;
using System.Threading.Tasks;

using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

using NLog.Web;

namespace InteractionOfficeBot.WebApi
{
	public class Program
	{
		public static async Task Main(string[] args)
		{
			var logger = NLogBuilder.ConfigureNLog("nlog.config").GetCurrentClassLogger();

			try
			{
				logger.Debug("init main");
				await CreateHostBuilder(args).Build().RunAsync();
			}
			catch (Exception exception)
			{
				//NLog: catch setup errors
				logger.Error(exception, "Stopped program because of exception");
				throw;
			}
			finally
			{
				// Ensure to flush and stop internal timers/threads before application-exit (Avoid segmentation fault on Linux)
				NLog.LogManager.Shutdown();
			}
		}

		public static IHostBuilder CreateHostBuilder(string[] args) =>
			Host.CreateDefaultBuilder(args)
				.ConfigureWebHostDefaults(webBuilder =>
				{
					webBuilder.ConfigureLogging((logging) =>
					{
						logging.AddDebug();
						logging.AddConsole();
					});
					webBuilder.UseStartup<Startup>();
				})
				.ConfigureLogging(logging =>
				{
					logging.ClearProviders();
					logging.SetMinimumLevel(LogLevel.Trace);
				}).UseNLog();
	}
}
