using System.Net.Http.Headers;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace InteractionOfficeBot.Core.MsGraph
{
	public interface IGraphServiceClientFactory
	{
		GraphServiceClient CreateClientFromUserBeHalf(string token);
		GraphServiceClient CreateClientFromApplicationBeHalf(IConfigurationRoot configuration);
	}

	public class GraphServiceClientFactory : IGraphServiceClientFactory
	{
		public GraphServiceClient CreateClientFromUserBeHalf(string token)
		{
			var graphClient = new GraphServiceClient(
				new DelegateAuthenticationProvider(
					requestMessage =>
					{
						// Append the access token to the request.
						requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);

						// Get event times in the current time zone.
						requestMessage.Headers.Add("Prefer", "outlook.timezone=\"" + TimeZoneInfo.Local.Id + "\"");

						return Task.CompletedTask;
					}));
			return graphClient;
		}

        public GraphServiceClient CreateClientFromApplicationBeHalf(IConfigurationRoot configuration)
		{
			var authenticationProvider = CreateAuthorizationProviderFromApplicationBeHalf(configuration);
			var graphClient = new GraphServiceClient(authenticationProvider);

			return graphClient;
		}

		private IAuthenticationProvider CreateAuthorizationProviderFromApplicationBeHalf(IConfigurationRoot config)
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
	}
}
