using InteractionOfficeBot.Core.MsGraph;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace InteractionOfficeBot.Console
{
	public class Program
	{
		public static async Task Main(string[] args)
		{
			var config = LoadAppSettings();
			if (config == null)
			{
				System.Console.WriteLine("Invalid appsettings.json file.");
				return;
			}
			var serviceProvider = CreateServiceProvider();

			var factory = serviceProvider.GetRequiredService<IGraphServiceClientFactory>();
			var client = factory.CreateClientFromApplicationBeHalf(config);

			var graphRequest = client.Users
				.Request()
				.Select(u => new { u.DisplayName, u.Mail });

			var results = await graphRequest.GetAsync();
			foreach (var user in results)
			{
				System.Console.WriteLine(user.Id + ": " + user.DisplayName + " <" + user.Mail + ">");
			}

			System.Console.WriteLine("\nGraph Request:");
			System.Console.WriteLine(graphRequest.GetHttpRequestMessage().RequestUri);
			System.Console.ReadKey();
		}

		private static IConfigurationRoot? LoadAppSettings()
		{
			try
			{
				var config = new ConfigurationBuilder()
								  .SetBasePath(System.IO.Directory.GetCurrentDirectory())
								  .AddJsonFile("appsettings.json", false, true)
								  .Build();

				if (string.IsNullOrEmpty(config["applicationId"]) ||
					string.IsNullOrEmpty(config["applicationSecret"]) ||
					string.IsNullOrEmpty(config["redirectUri"]) ||
					string.IsNullOrEmpty(config["tenantId"]))
				{
					return null;
				}

				return config;
			}
			catch (System.IO.FileNotFoundException)
			{
				return null;
			}
		}

		private static IServiceProvider CreateServiceProvider()
		{
			var services = new ServiceCollection();

			services.AddSingleton<IGraphServiceClientFactory, GraphServiceClientFactory>();

			return services.BuildServiceProvider();
		}
	}
}