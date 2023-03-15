using AdaptiveCards.Templating;
using InteractionOfficeBot.Core.Exception;
using InteractionOfficeBot.Core.Model;
using InteractionOfficeBot.Core.MsGraph;
using Microsoft.Bot.Builder;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Directory = System.IO.Directory;
using Microsoft.Identity.Client;
using System.Threading;


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

            //var a = await client.Teams.GetMembersOfTeams("Mark 8 Project Team");

            //foreach (var user in a)
            //{
	           // var userInfo = user.DisplayName + " Activity:" + user.Activity;
	           // System.Console.WriteLine(userInfo);
            //}

            IEnumerable<TeamsUserInfo> users;
            try
            {
	            users = await client.GetUsers();
            }
            catch (TeamsException e)
            {
	            //await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
	            return;
            }

            var path = Path.Combine(Environment.CurrentDirectory, @"AdaptiveCard", "UserCard.json");
            
            var templateJson = await File.ReadAllTextAsync(path);
            var template = new AdaptiveCardTemplate(templateJson);

            foreach (var user in users)
            {
	            var cardJson = template.Expand(user);
            }

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