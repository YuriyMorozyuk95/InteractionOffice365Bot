using InteractionOfficeBot.Core.MsGraph;

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace InteractionOfficeBot.Console
{
	public class Program
	{
		public static async Task Main(string[] args)
		{
			var serviceProvider = CreateServiceProvider();

			var factory = serviceProvider.GetRequiredService<IGraphServiceClientFactory>();
			var client = factory.CreateClientFromApplicationBeHalf();

			var users =  await client.GetUsers();

			foreach (var user in users)
			{
				System.Console.WriteLine(user.Id + ": " + user.DisplayName + " <" + user.Mail + ">");
			}

			System.Console.ReadKey();
		}

		private static IConfigurationRoot LoadAppSettings()
		{
			var config = new ConfigurationBuilder()
								  .SetBasePath(Directory.GetCurrentDirectory())
								  .AddJsonFile("appsettings.json", false, true)
								  .Build();

			if (string.IsNullOrEmpty(config["applicationId"]) ||
				string.IsNullOrEmpty(config["applicationSecret"]) ||
				string.IsNullOrEmpty(config["redirectUri"]) ||
				string.IsNullOrEmpty(config["tenantId"]))
			{
				throw new Exception("Missing app registration properties");
			}

			return config;
		}

		private static IServiceProvider CreateServiceProvider()
		{
			var config = LoadAppSettings();

			var services = new ServiceCollection();

			services.AddSingleton<IConfiguration>(config);
			services.AddSingleton<IGraphServiceClientFactory, GraphServiceClientFactory>();

			return services.BuildServiceProvider();
		}
	}
}