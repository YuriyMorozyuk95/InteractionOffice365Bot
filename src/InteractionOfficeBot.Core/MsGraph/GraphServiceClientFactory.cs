using System.Net.Http.Headers;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace InteractionOfficeBot.Core.MsGraph
{
	public interface IGraphServiceClientFactory
	{
		IobGraphClient CreateClientFromUserBeHalf(string token);
		IobGraphClient CreateClientFromApplicationBeHalf();
	}

	public class GraphServiceClientFactory : IGraphServiceClientFactory
	{
		private readonly IConfiguration _configuration;

		public GraphServiceClientFactory(IConfiguration configuration)
		{
			_configuration = configuration;
		}
        public IobGraphClient CreateClientFromUserBeHalf(string token)
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
			return new IobGraphClient(graphClient);
		}

        public IobGraphClient CreateClientFromApplicationBeHalf()
		{
			var authenticationProvider = CreateAuthorizationProviderFromApplicationBeHalf();
			var graphClient = new GraphServiceClient(authenticationProvider);

			return new IobGraphClient(graphClient);
		}

		private IAuthenticationProvider CreateAuthorizationProviderFromApplicationBeHalf()
		{
			var clientId = _configuration["applicationId"];
			var clientSecret = _configuration["applicationSecret"];
			var redirectUri = _configuration["redirectUri"];
			var authority = $"https://login.microsoftonline.com/{_configuration["tenantId"]}/v2.0";

			var scopes = new List<string> { "https://graph.microsoft.com/.default" };

			var cca = ConfidentialClientApplicationBuilder.Create(clientId)
				.WithAuthority(authority)
				.WithRedirectUri(redirectUri)
				.WithClientSecret(clientSecret)
				.Build();
			return new MsalAuthenticationProvider(cca, scopes.ToArray());
		}
	}
}
