using InteractionOfficeBot.Console.Helpers;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;

namespace InteractionOfficeBot.Console
{
	public class Program
	{
		private static GraphServiceClient? _graphClient;

		public static async Task Main(string[] args)
		{
			System.Console.WriteLine("Hello World!");

			var config = LoadAppSettings();
			if (config == null)
			{
				System.Console.WriteLine("Invalid appsettings.json file.");
				return;
			}

			var client = GetAuthenticatedGraphClient(config);

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

		private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
		{
			var clientId = config["applicationId"];
			var clientSecret = config["applicationSecret"];
			var redirectUri = config["redirectUri"];
			var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

			List<string> scopes = new List<string>();
			scopes.Add("https://graph.microsoft.com/.default");

			var cca = ConfidentialClientApplicationBuilder.Create(clientId)
													.WithAuthority(authority)
													.WithRedirectUri(redirectUri)
													.WithClientSecret(clientSecret)
													.Build();
			return new MsalAuthenticationProvider(cca, scopes.ToArray());
		}

		private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
		{
			var authenticationProvider = CreateAuthorizationProvider(config);
			_graphClient = new GraphServiceClient(authenticationProvider);
			return _graphClient;
		}
	}
}