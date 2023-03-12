using System;
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
	    private const string ALL_USER_REQUEST = "Show me all users";
	    private const string ALL_TEAMS_REQUEST = "Show all teams in organization";
	    private const string WHO_OF_TEAMS_REQUEST = "Who is on the team: 'Test Team'?";
	    private const string WHAT_CHANNELS_OF_TEAMS_REQUEST = "What channels can I find in the team:'Test Team'?";
	    private const string CREATE_TEAM = "Please create team:'Test Team' for user:'victoria@8bpskq.onmicrosoft.com'.";
	    private const string CREATE_CHANNEL = "Please create chanel: 'Test Chanel' for team:'Test Team'.";
	    private const string MEMBER_CHANNEL = "Who are members of chanel: 'Test Chanel' in team: 'Test Team'?";
	    private const string REMOVE_CHANNEL = "Please remove chanel: 'Test Chanel' in team: 'Test Team'";
	    private const string REMOVE_TEAM = "Please remove team: 'Test Team'";
	    private const string SEND_MESSAGE_TO_CHANEL = "Please send message: 'Hello world' to channel: 'Test Chanel' in team: 'Test Team'";

	    private const string GraphDialog = "GraphDialog";



	    private readonly ILogger _logger;
        private readonly IStateService _stateService;
        private readonly IGraphServiceClientFactory _graphServiceClient;
        private readonly int _expireAfterMinutes;

        public MainDialog(IConfiguration configuration, ILogger<MainDialog> logger, IStateService stateService, IGraphServiceClientFactory graphServiceClient)
            : base(nameof(MainDialog), configuration["ConnectionName"])
        {
            _logger = logger;
            _stateService = stateService;
            _graphServiceClient = graphServiceClient;
            _expireAfterMinutes = configuration.GetValue<int>("ExpireAfterMinutes");

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

            AddDialog(new TextPrompt(GraphDialog));

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
	        // Retrieve the property value, and compare it to the current time.
	        var lastAccess = await _stateService.LastAccessedTimeAccessor.GetAsync(
		        stepContext.Context,
		        () => DateTime.UtcNow,
		        cancellationToken)
		        .ConfigureAwait(false);

	        if ((DateTime.UtcNow - lastAccess) >= TimeSpan.FromMinutes(_expireAfterMinutes))
	        {
		        // Notify the user that the conversation is being restarted.
		        await stepContext.Context.SendActivityAsync("Welcome back!  Let's start over from the beginning.").ConfigureAwait(false);

		        // Clear state.
		        await _stateService.ConversationState.ClearStateAsync(stepContext.Context, cancellationToken).ConfigureAwait(false);
		        await _stateService.UserState.ClearStateAsync(stepContext.Context, cancellationToken).ConfigureAwait(false);

		        return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken);
	        }

	        return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken);
        }

        private async Task<DialogTurnResult> LoginStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
	        var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);

	        if (userTokeStore.Token != null)
	        {
		        return await stepContext.PromptAsync(
			        GraphDialog,
			        new PromptOptions { Prompt = MessageFactory.Text("Can I help you with something else?") },
			        cancellationToken);
	        }

            // Get the token from the previous step. Note that we could also have gotten the
            // token directly from the prompt itself. There is an example of this in the next method.
            var tokenResponse = (TokenResponse)stepContext.Result;
            if (tokenResponse?.Token != null)
            {
				userTokeStore.Token = tokenResponse.Token;
	            await _stateService.UserTokeStoreAccessor.SetAsync(stepContext.Context, userTokeStore, cancellationToken);

                // Pull in the data from the Microsoft Graph.
                //TODO Try Put GraphClient to UserTokenStore
                var client =  _graphServiceClient.CreateClientFromUserBeHalf(tokenResponse.Token);
                var me = await client.GetMeAsync();
                var title = !string.IsNullOrEmpty(me.JobTitle) ?
                            me.JobTitle : "Unknown";

				await stepContext.Context.SendActivityAsync($"You're logged in as {me.DisplayName} ({me.UserPrincipalName}); you job title is: {title}", cancellationToken: cancellationToken);

                return await stepContext.PromptAsync(GraphDialog, new PromptOptions { Prompt = MessageFactory.Text("How I can help you?") }, cancellationToken);
            }

            await stepContext.Context.SendActivityAsync(MessageFactory.Text("Login was not successful please try again."), cancellationToken);
            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
        }

		private async Task<DialogTurnResult> GraphActionStep(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var result = (string)stepContext.Result;

			switch (result)
			{
				case ALL_USER_REQUEST:
					await ShowAllUsers(stepContext, cancellationToken);
					break;
				case ALL_TEAMS_REQUEST:
					await ShowAllTeams(stepContext, cancellationToken);
					break;
				case WHO_OF_TEAMS_REQUEST:
					await MemeberOfTeam(stepContext, cancellationToken, "Test Team");
					break;
				case WHAT_CHANNELS_OF_TEAMS_REQUEST:
					await ChanelOfTeam(stepContext, cancellationToken, "Test Team");
					break;
				case CREATE_TEAM:
					await CreateTeamFor(stepContext, cancellationToken, "Test Team", "victoria@8bpskq.onmicrosoft.com");
					break;
				case CREATE_CHANNEL:
					await CreateChanelForTeam(stepContext, cancellationToken, "Test Team", "Test Chanel");
					break;
				case MEMBER_CHANNEL:
					await MemberOfChanel(stepContext, cancellationToken, "Test Team", "Test Chanel");
					break;
				case REMOVE_CHANNEL:
					await RemoveChanel(stepContext, cancellationToken, "Test Team", "Test Chanel");
					break;
				case REMOVE_TEAM:
					await RemoveTeam(stepContext, cancellationToken, "Test Team");
					break;
				case SEND_MESSAGE_TO_CHANEL:
					await SendMessageToChanel(stepContext, cancellationToken, "Test Team", "Test Chanel", "Hello, world");
					break;
			}
			//return await stepContext.ReplaceDialogAsync(InitialDialogId, new object(), cancellationToken);
			return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
		}

		private async Task SendMessageToChanel(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string channelName, string message)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			await client.Teams.SendMessageToChanel(teamName, channelName, message);

			await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Message was send"), cancellationToken);
		}

		private async Task RemoveTeam(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			await client.Teams.RemoveTeam(teamName);

			await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Team with name: {teamName} was removed"), cancellationToken);
		}

		private async Task RemoveChanel(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string channelName)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			await client.Teams.RemoveChannelFromTeam(teamName, channelName);

			await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Channel with name: {channelName} for team: {teamName} was removed"), cancellationToken);
		}

		private async Task MemberOfChanel(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string channelName)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			var users = await client.Teams.GetMembersOfChannelFromTeam(teamName, channelName);

			foreach (var user in users)
			{
				var userInfo = user.DisplayName;
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(userInfo), cancellationToken);
			}
		}

		private async Task CreateChanelForTeam(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string channelName)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			await client.Teams.CreateChannelForTeam(teamName, channelName);

			await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Channel with name: {channelName} for team: {teamName} was created"), cancellationToken);
		}

		private async Task CreateTeamFor(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string userEmail)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			await client.Teams.CreateTeamFor(teamName, userEmail);

			await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Team with name: {teamName} for {userEmail} was created"), cancellationToken);
		}

		private async Task ChanelOfTeam(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			var channels = await client.Teams.GetChannelsOfTeams(teamName);

			foreach (var channel in channels)
			{
				var chanelInfo = channel.DisplayName + " Link: "+ channel.WebUrl;
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(chanelInfo), cancellationToken);
			}
		}

		private async Task MemeberOfTeam(WaterfallStepContext stepContext, CancellationToken cancellationToken, string testTeam)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			var users = await client.Teams.GetMembersOfTeams(testTeam);

			foreach (var user in users)
			{
				var userInfo = user.DisplayName;
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(userInfo), cancellationToken);
			}
		}

		private async Task ShowAllTeams(WaterfallStepContext stepContext, CancellationToken cancellationToken)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			var teams = await client.Teams.GetListTeams();

			foreach (var team in teams)
			{
				var teamInfo = team.DisplayName + "Description: " + team.Description;
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(teamInfo), cancellationToken);
			}
		}

		private async Task ShowAllUsers(WaterfallStepContext stepContext, CancellationToken cancellationToken)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			var users = await client.GetUsers();

			foreach (var user in users)
			{
				var userInfo = user.DisplayName + " <" + user.Mail + ">";
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(userInfo), cancellationToken);
			}
		}
    }
}
