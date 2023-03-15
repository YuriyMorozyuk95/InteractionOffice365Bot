using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using InteractionOfficeBot.Core.Exception;
using InteractionOfficeBot.Core.Model;
using InteractionOfficeBot.Core.MsGraph;
using InteractionOfficeBot.WebApi.Helper;
using InteractionOfficeBot.WebApi.Services;

using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
namespace InteractionOfficeBot.WebApi.Dialogs
{
	public class MainDialog : LogoutDialog
	{
		private const string ALL_USER_REQUEST = "Show me all users";
		private const string ALL_TEAMS_REQUEST = "Show all teams in organization";
		private const string WHO_OF_TEAMS_REQUEST = "Who is on the team: 'Test Team'?";
		private const string WHAT_CHANNELS_OF_TEAMS_REQUEST = "What channels can I find in the team:'Test Team'?";
		private const string CREATE_TEAM = "Please create team:'Test Team' for user:'victoria@8bpskq.onmicrosoft.com'.";
		private const string CREATE_CHANNEL = "Please create channel: 'Test Chanel' for team:'Test Team'.";
		private const string MEMBER_CHANNEL = "Who are members of chanel: 'Test Chanel' in team: 'Test Team'?";
		private const string REMOVE_CHANNEL = "Please remove chanel: 'Test Chanel' in team: 'Test Team'";
		private const string REMOVE_TEAM = "Please remove team: 'Test Team'";
		private const string SEND_MESSAGE_TO_CHANEL = "Please send message: 'Hello world' to channel: 'Test Chanel' in team: 'Test Team'";
		private const string SEND_EMAIL_TO_USER = "Please send email with subject: 'test' and message: 'Hello world' to user: 'victoria@8bpskq.onmicrosoft.com'";
		private const string INSTALLED_APP_FOR_USER = "Show me all installed applications in teams for user: 'victoria@8bpskq.onmicrosoft.com'";

		//TODO ADD to LUIS
		//TODO ADD entity email to LUIS
		private const string ONEDRIVE_ROOT_CONTENTS = "Show me files in OneDrive";
		private const string ONEDRIVE_FOLDER_CONTENTS = "Show me files in folder : 'vika dura/alex loh'";
		private const string ONEDRIVE_SEARCH = "Search for files : 'paris'";
		private const string ONEDRIVE_DELETE = "Remove file : 'vika dura/my-picture.jpeg'";
		private const string ONEDRIVE_DOWNLOAD = "Download file : 'vika dura/my-picture.jpeg;";

		private const string GraphDialog = "GraphDialog";

		private readonly ILogger _logger;
		private readonly IStateService _stateService;
		private readonly IGraphServiceClientFactory _graphServiceClient;
		private readonly ILuisService _luisService;
		private readonly IGraphDialogHelper _graphDialogHelper;
		private readonly int _expireAfterMinutes;

		public MainDialog(
			IConfiguration configuration,
			ILogger<MainDialog> logger,
			IStateService stateService,
			IGraphServiceClientFactory graphServiceClient,
			ILuisService luisService,
			IGraphDialogHelper graphDialogHelper)
			: base(nameof(MainDialog), configuration["ConnectionName"])
		{
			_logger = logger;
			_stateService = stateService;
			_graphServiceClient = graphServiceClient;
			_luisService = luisService;
			_graphDialogHelper = graphDialogHelper;
			_expireAfterMinutes = configuration.GetValue<int>("ExpireAfterMinutes");

			AddDialog(new OAuthPrompt(
				nameof(OAuthPrompt),
				new OAuthPromptSettings
				{
					ConnectionName = ConnectionName,
					Text = "Please Sign In",
					Title = "Sign In",
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
				_logger.LogWarning("token expired");

				// Notify the user that the conversation is being restarted.
				await stepContext.Context.SendActivityAsync("Welcome back!  Let's start over from the beginning.").ConfigureAwait(false);

				// Clear state.
				await _stateService.ConversationState.ClearStateAsync(stepContext.Context, cancellationToken).ConfigureAwait(false);
				await _stateService.UserState.ClearStateAsync(stepContext.Context, cancellationToken).ConfigureAwait(false);

				return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken);
			}

			_logger.LogWarning("token still valid");

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
				var client = _graphServiceClient.CreateClientFromUserBeHalf(tokenResponse.Token);
				var me = await client.GetMeAsync();

				await stepContext.Context.SendActivityAsync($"You're logged in as {me.DisplayName} ({me.UserPrincipalName})", cancellationToken: cancellationToken);

				return await stepContext.PromptAsync(GraphDialog, new PromptOptions { Prompt = MessageFactory.Text("How I can help you?") }, cancellationToken);
			}

			await stepContext.Context.SendActivityAsync(MessageFactory.Text("Login was not successful please try again."), cancellationToken);
			return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
		}

		private async Task<DialogTurnResult> GraphActionStep(WaterfallStepContext stepContext, CancellationToken cancellationToken)
		{
			var recognizeResult = await _luisService.Recognize(stepContext.Context, cancellationToken);
			var topIntent = recognizeResult.TopIntent();

			switch (topIntent.intent)
			{
				case LuisRoot.Intent.ALL_USER_REQUEST:
					await _graphDialogHelper.ShowAllUsers(stepContext, cancellationToken);
					break;
				case LuisRoot.Intent.ALL_TEAMS_REQUEST:
					await _graphDialogHelper.ShowAllTeams(stepContext, cancellationToken);
					break;
				case LuisRoot.Intent.WHO_OF_TEAMS_REQUEST:
					await _graphDialogHelper.MembersOfTeam(stepContext, cancellationToken, LuisEntityHelper.GetTeamFromEntity(recognizeResult));
					break;
				case LuisRoot.Intent.WHAT_CHANNELS_OF_TEAMS_REQUEST:
					await _graphDialogHelper.ChanelOfTeam(stepContext, cancellationToken, LuisEntityHelper.GetTeamFromEntity(recognizeResult));
					break;
				case LuisRoot.Intent.CREATE_TEAM:
					await _graphDialogHelper.CreateTeamFor(stepContext, cancellationToken, LuisEntityHelper.GetTeamFromEntity(recognizeResult), LuisEntityHelper.GetUserFromEntity(recognizeResult));
					break;
				case LuisRoot.Intent.CREATE_CHANNEL:
					await _graphDialogHelper.CreateChanelForTeam(stepContext, cancellationToken, LuisEntityHelper.GetTeamFromEntity(recognizeResult), LuisEntityHelper.GetChannelFromEntity(recognizeResult));
					break;
				case LuisRoot.Intent.MEMBER_CHANNEL:
					await _graphDialogHelper.MemberOfChanel(stepContext, cancellationToken, LuisEntityHelper.GetTeamFromEntity(recognizeResult), LuisEntityHelper.GetChannelFromEntity(recognizeResult));
					break;
				case LuisRoot.Intent.REMOVE_CHANNEL:
					await _graphDialogHelper.RemoveChanel(stepContext, cancellationToken, LuisEntityHelper.GetTeamFromEntity(recognizeResult), LuisEntityHelper.GetChannelFromEntity(recognizeResult));
					break;
				case LuisRoot.Intent.REMOVE_TEAM:
					await _graphDialogHelper.RemoveTeam(stepContext, cancellationToken, LuisEntityHelper.GetTeamFromEntity(recognizeResult));
					break;
				case LuisRoot.Intent.SEND_MESSAGE_TO_CHANEL:
					await _graphDialogHelper.SendMessageToChanel(stepContext, cancellationToken, LuisEntityHelper.GetTeamFromEntity(recognizeResult), LuisEntityHelper.GetChannelFromEntity(recognizeResult), LuisEntityHelper.GetMessageFromEntity(recognizeResult));
					break;
				case LuisRoot.Intent.SEND_EMAIL_TO_USER:
					await _graphDialogHelper.SendEmailToUser(stepContext, cancellationToken, LuisEntityHelper.GetUserFromEntity(recognizeResult), LuisEntityHelper.GetEmailSubjectFromEntity(recognizeResult), LuisEntityHelper.GetMessageFromEntity(recognizeResult));
					break;
				case LuisRoot.Intent.INSTALLED_APP_FOR_USER:
					await _graphDialogHelper.ShowInstalledAppForUser(stepContext, cancellationToken, LuisEntityHelper.GetUserFromEntity(recognizeResult));
					break;
				case LuisRoot.Intent.ONEDRIVE_ROOT_CONTENTS:
					await _graphDialogHelper.ShowOneDriveContents(stepContext, cancellationToken);
					break;
				case LuisRoot.Intent.ONEDRIVE_FOLDER_CONTENTS:
					await _graphDialogHelper.ShowOneDriveFolderContents(stepContext, cancellationToken, LuisEntityHelper.GetFolderPathFromEntity(recognizeResult));
					break;
				case LuisRoot.Intent.ONEDRIVE_SEARCH:
					await _graphDialogHelper.SearchOneDrive(stepContext, cancellationToken, LuisEntityHelper.GetFileFromEntity(recognizeResult));
					break;
				case LuisRoot.Intent.ONEDRIVE_DELETE:
					await _graphDialogHelper.DeleteOneDrive(stepContext, cancellationToken, LuisEntityHelper.GetFileFromEntity(recognizeResult));
					break;
				case LuisRoot.Intent.ONEDRIVE_DOWNLOAD:
					await _graphDialogHelper.DownloadOneDrive(stepContext, cancellationToken, LuisEntityHelper.GetFileFromEntity(recognizeResult));
					break;
				case LuisRoot.Intent.GET_ALL_TODO_TASKS:
					await _graphDialogHelper.GetAllTodoTasks(stepContext, cancellationToken);
					break;
				case LuisRoot.Intent.GET_ALL_TODO_UPCOMING_TASK:
					await _graphDialogHelper.GetTodoUpcomingTask(stepContext, cancellationToken, LuisEntityHelper.GetTaskReminderTimeFromEntity(recognizeResult));
					break;
				case LuisRoot.Intent.CREATE_TODO_TASK:
					await _graphDialogHelper.CreateTodoTask(stepContext, cancellationToken, LuisEntityHelper.GetTaskTitleFromEntity(recognizeResult), LuisEntityHelper.GetTaskReminderTimeFromEntity(recognizeResult));
					break;
			}
			return await stepContext.BeginDialogAsync(InitialDialogId, null, cancellationToken);
		}
	}
}
