using Microsoft.Graph;

namespace InteractionOfficeBot.Console.Helpers
{
	public class AuthHandler : DelegatingHandler
	{
		private readonly IAuthenticationProvider _authenticationProvider;

		public AuthHandler(IAuthenticationProvider authenticationProvider, HttpMessageHandler innerHandler)
		{
			InnerHandler = innerHandler;
			_authenticationProvider = authenticationProvider;
		}

		protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
		{
			await _authenticationProvider.AuthenticateRequestAsync(request);
			return await base.SendAsync(request, cancellationToken);
		}
	}
}
