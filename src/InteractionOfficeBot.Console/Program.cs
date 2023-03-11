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

			var teams = await client.Client.Teams.Request().Filter("startswith(displayName,%20'A')").Top(1).GetAsync();

			System.Console.ReadKey();
		}


		private static async Task CreateTeams(IobGraphClient client)
		{
			await client.TeamsGroup.CreateTeamFor("test1", "yurii.moroziuk.iob@8bpskq.onmicrosoft.com");
		}

		private static async Task WriteTeamsList(IobGraphClient client)
		{
			var teams = await client.TeamsGroup.GetListTeams();

			foreach (var team in teams)
			{
				System.Console.WriteLine(team.Id + ": " + team.DisplayName);
			}
		}

		private static async Task WriteLineUserList(IobGraphClient client)
		{
			var users = await client.GetUsers();

			foreach (var user in users)
			{
				System.Console.WriteLine(user.Id + ": " + user.DisplayName + " <" + user.Mail + ">");
			}
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