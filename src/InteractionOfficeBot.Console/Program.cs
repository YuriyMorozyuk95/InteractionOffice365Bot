using InteractionOfficeBot.Core.MsGraph;

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Directory = System.IO.Directory;
using Microsoft.Identity.Client;


namespace InteractionOfficeBot.Console
{
	public class Program
	{
		public static async Task Main(string[] args)
		{
			var serviceProvider = CreateServiceProvider();

			var factory = serviceProvider.GetRequiredService<IGraphServiceClientFactory>();

			var scopes = new [] { "https://graph.microsoft.com/.default" };
			var token = await GetUserToken(scopes);

            var client = factory.CreateClientFromUserBeHalf(token);

			//await client.SendMailAsync("yurii.moroziuk.iob@8bpskq.onmicrosoft.com", "hello", "hey");
			await client.SendMailAsync("yurii.moroziuk@hotmail.com", "hello", "hey");

            System.Console.ReadKey();
		}


		private static async Task CreateTeams(IobGraphClient client)
		{
			await client.Teams.CreateTeamFor("test1", "yurii.moroziuk.iob@8bpskq.onmicrosoft.com");
		}

		private static async Task WriteTeamsList(IobGraphClient client)
		{
			var teams = await client.Teams.GetListTeams();

			foreach (var team in teams)
			{
				System.Console.WriteLine(team.Id + ": " + team.DisplayName);
			}
		}

		private static async Task WriteChannelList(IobGraphClient client)
		{
			var teams = await client.Teams.GetChannelsOfTeams("test1");

			foreach (var team in teams)
			{
				System.Console.WriteLine(team.DisplayName + ": " + team.WebUrl);
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
				string.IsNullOrEmpty(config["tenantId"]) ||
				string.IsNullOrEmpty(config["authority"]))
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

        private static async Task<string> GetUserToken(string[] scopes)
        {
            var config = LoadAppSettings();

            var _publicClientApplication = PublicClientApplicationBuilder.Create(config["applicationId"])
				.WithAuthority(config["authority"])
				.WithDefaultRedirectUri()
				.Build();

            var _authResult = await _publicClientApplication.AcquireTokenInteractive(scopes).ExecuteAsync();

            return _authResult.AccessToken;

        }
    }
}