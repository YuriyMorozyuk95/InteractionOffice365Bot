using System.Threading;
using System.Threading.Tasks;
using InteractionOfficeBot.Core.MsGraph;
using InteractionOfficeBot.WebApi.Services;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace InteractionOfficeBot.WebApi.Dialogs
{
    public class MainDialog : LogoutDialog
    {
	    private const string ALL_USER_REQUEST = "show me all users";

	    private readonly ILogger _logger;
        private readonly IStateService _stateService;
        private readonly IGraphServiceClientFactory _graphServiceClient;

        public MainDialog(IConfiguration configuration, ILogger<MainDialog> logger, IStateService stateService, IGraphServiceClientFactory graphServiceClient)
            : base(nameof(MainDialog), configuration["ConnectionName"])
        {
            _logger = logger;
            _stateService = stateService;
            _graphServiceClient = graphServiceClient;

            AddDialog(new OAuthPrompt(
                nameof(OAuthPrompt),
                new OAuthPromptSettings
                {
                    ConnectionName = ConnectionName,
                    Text = "Please Sign In",
                    Title = "Sign In",
                    Timeout = 300000, // User has 5 minutes to login (1000 * 60 * 5)
                    EndOnInvalidMessage = true
                }));

            AddDialog(new TextPrompt(nameof(TextPrompt)));

            AddDialog(new WaterfallDialog(nameof(WaterfallDialog), new WaterfallStep[]
            {
                PromptStepAsync,
                LoginStepAsync,
                GraphActionStep,
            }));

            // The initial child Dialog to run.
            InitialDialogId = nameof(WaterfallDialog);
        }

        private async Task<DialogTurnResult> PromptStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken);
        }

        private async Task<DialogTurnResult> LoginStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Get the token from the previous step. Note that we could also have gotten the
            // token directly from the prompt itself. There is an example of this in the next method.
            var tokenResponse = (TokenResponse)stepContext.Result;
            if (tokenResponse?.Token != null)
            {
	            var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
	            userTokeStore.Token = tokenResponse.Token;
	            await _stateService.UserTokeStoreAccessor.SetAsync(stepContext.Context, userTokeStore, cancellationToken);


                // Pull in the data from the Microsoft Graph.
                //TODO Try Put GraphClient to UserTokenStore
                var client =  _graphServiceClient.CreateClientFromUserBeHalf(tokenResponse.Token);
                var me = await client.GetMeAsync();
                var title = !string.IsNullOrEmpty(me.JobTitle) ?
                            me.JobTitle : "Unknown";

				await stepContext.Context.SendActivityAsync($"You're logged in as {me.DisplayName} ({me.UserPrincipalName}); you job title is: {title}", cancellationToken: cancellationToken);

                return await stepContext.PromptAsync(nameof(TextPrompt), new PromptOptions { Prompt = MessageFactory.Text("How I can help you?") }, cancellationToken);
            }

            await stepContext.Context.SendActivityAsync(MessageFactory.Text("Login was not successful please try again."), cancellationToken);
            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
        }


		private async Task<DialogTurnResult> GraphActionStep(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var result = (string)stepContext.Result;

            //TODO add const
            if (result == ALL_USER_REQUEST)
            {
	            var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
	            var client =  _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

	           var users = await client.GetUsers();

	           foreach (var user in users)
	           {
		           var userInfo = user.Id + ": " + user.DisplayName + " <" + user.Mail + ">";
		           await stepContext.Context.SendActivityAsync(MessageFactory.Text(userInfo), cancellationToken);
	           }

            }

            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
        }
    }
}
