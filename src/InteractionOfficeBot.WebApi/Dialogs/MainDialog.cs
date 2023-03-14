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
using InteractionOfficeBot.WebApi.Services;

using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

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
		private readonly int _expireAfterMinutes;

		public MainDialog(
			IConfiguration configuration,
			ILogger<MainDialog> logger,
			IStateService stateService,
			IGraphServiceClientFactory graphServiceClient,
			ILuisService luisService)
			: base(nameof(MainDialog), configuration["ConnectionName"])
		{
			_logger = logger;
			_stateService = stateService;
			_graphServiceClient = graphServiceClient;
			_luisService = luisService;
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
				//TODO Try Put GraphClient to UserTokenStore
				var client = _graphServiceClient.CreateClientFromUserBeHalf(tokenResponse.Token);
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

			var recognizeResult = await _luisService.Recognize(stepContext.Context, cancellationToken);
			var topIntent = recognizeResult.TopIntent();

			switch (result)
			{
				default:
					switch (topIntent.intent)
					{
						case LuisRoot.Intent.ALL_USER_REQUEST:
							await ShowAllUsers(stepContext, cancellationToken);
							break;
						case LuisRoot.Intent.ALL_TEAMS_REQUEST:
							await ShowAllTeams(stepContext, cancellationToken);
							break;
						case LuisRoot.Intent.WHO_OF_TEAMS_REQUEST:
							{
								await MemeberOfTeam(
									stepContext,
									cancellationToken,
									GetTeamFromEntity(recognizeResult));
								break;
							}
						case LuisRoot.Intent.WHAT_CHANNELS_OF_TEAMS_REQUEST:
							await ChanelOfTeam(stepContext, cancellationToken, GetTeamFromEntity(recognizeResult));
							break;
						case LuisRoot.Intent.CREATE_TEAM:
							await CreateTeamFor(stepContext, cancellationToken, GetTeamFromEntity(recognizeResult), GetChannelFromEntity(recognizeResult));
							break;
						case LuisRoot.Intent.CREATE_CHANNEL:
							await CreateChanelForTeam(stepContext, cancellationToken, GetTeamFromEntity(recognizeResult), GetChannelFromEntity(recognizeResult));
							break;
						case LuisRoot.Intent.MEMBER_CHANNEL:
							await MemberOfChanel(stepContext, cancellationToken, GetTeamFromEntity(recognizeResult), GetChannelFromEntity(recognizeResult));
							break;
						case LuisRoot.Intent.REMOVE_CHANNEL:
							await RemoveChanel(stepContext, cancellationToken, GetTeamFromEntity(recognizeResult), GetChannelFromEntity(recognizeResult));
							break;
						case LuisRoot.Intent.REMOVE_TEAM:
							await RemoveTeam(stepContext, cancellationToken, GetTeamFromEntity(recognizeResult));
							break;
						case LuisRoot.Intent.SEND_MESSAGE_TO_CHANEL:
							await SendMessageToChanel(stepContext, cancellationToken, GetTeamFromEntity(recognizeResult), GetChannelFromEntity(recognizeResult), GetMessageFromEntity(recognizeResult));
							break;
						case LuisRoot.Intent.SEND_EMAIL_TO_USER:
							await SendEmailToUser(stepContext, cancellationToken, GetUserFromEntity(recognizeResult), GetEmailSubjectFromEntity(recognizeResult), GetMessageFromEntity(recognizeResult));
							break;
						case LuisRoot.Intent.INSTALLED_APP_FOR_USER:
							await ShowInstalledAppForUser(stepContext, cancellationToken, GetUserFromEntity(recognizeResult));
							break;
						case LuisRoot.Intent.ONEDRIVE_ROOT_CONTENTS:
							await ShowOneDriveContents(stepContext, cancellationToken);
							break;
						case LuisRoot.Intent.ONEDRIVE_FOLDER_CONTENTS:
							await ShowOneDriveFolderContents(stepContext, cancellationToken, GetFolderPathFromEntity(recognizeResult) );
							break;
						case LuisRoot.Intent.ONEDRIVE_SEARCH:
							await SearchOneDrive(stepContext, cancellationToken, GetFileFromEntity(recognizeResult));
							break;
						case LuisRoot.Intent.ONEDRIVE_DELETE:
							await DeleteOneDrive(stepContext, cancellationToken, GetFileFromEntity(recognizeResult));
							break;
						case LuisRoot.Intent.ONEDRIVE_DOWNLOAD:
							await DownloadOneDrive(stepContext, cancellationToken, GetFileFromEntity(recognizeResult));
							break;
					}
					break;
			}
			await stepContext.Context.SendActivityAsync(MessageFactory.Text("type something to continue"), cancellationToken);
			return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
		}

		#region Helpers

		private static string GetFileFromEntity(LuisRoot recognizeResult)
		{
			var team = recognizeResult.Entities
				?.Files
				?.FirstOrDefault()
				?.Value
				?.FirstOrDefault();

			if (team == null)
			{
				throw new TeamsException("Can't recognize file");
			}

			return team;
		}

		private static string GetFolderPathFromEntity(LuisRoot recognizeResult)
		{
			var team = recognizeResult.Entities
				?.Folder
				?.FirstOrDefault()
				?.Value
				?.FirstOrDefault();

			if (team == null)
			{
				throw new TeamsException("Can't recognize folder");
			}

			return team;
		}

		private static string GetTeamFromEntity(LuisRoot recognizeResult)
		{
			var team = recognizeResult.Entities
				?.Team
				?.FirstOrDefault()
				?.Value
				?.FirstOrDefault();

			if (team == null)
			{
				throw new TeamsException("Can't recognize team");
			}

			return team;
		}

		private static string GetChannelFromEntity(LuisRoot recognizeResult)
		{
			var channel = recognizeResult.Entities
				?.Channel
				?.FirstOrDefault()
				?.Value
				?.FirstOrDefault();

			if (channel == null)
			{
				throw new TeamsException("Can't recognize channel");
			}

			return channel;
		}

		private static string GetUserFromEntity(LuisRoot recognizeResult)
		{
			var channel = recognizeResult.Entities
				?.User
				?.FirstOrDefault()
				?.Value
				?.FirstOrDefault();

			if (channel == null)
			{
				throw new TeamsException("Can't recognize user email");
			}

			return channel;
		}

		private static string GetMessageFromEntity(LuisRoot recognizeResult)
		{
			var channel = recognizeResult.Entities
				?.Message
				?.FirstOrDefault()
				?.Value
				?.FirstOrDefault();

			if (channel == null)
			{
				throw new TeamsException("Can't recognize message");
			}

			return channel;
		}

		private static string GetEmailSubjectFromEntity(LuisRoot recognizeResult)
		{
			var channel = recognizeResult.Entities
				?.EmailSubject
				?.FirstOrDefault()
				?.Value
				?.FirstOrDefault();

			if (channel == null)
			{
				throw new TeamsException("Can't recognize message");
			}

			return channel;
		}

		#endregion 


		#region GraphMessageHandlers
		private async Task ShowInstalledAppForUser(WaterfallStepContext stepContext, CancellationToken cancellationToken, string email)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			IUserTeamworkInstalledAppsCollectionPage apps;
			try
			{
				apps = await client.Teams.GetInstalledAppForUser(email);
			}
			catch (TeamsException e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

			foreach (var channel in apps)
			{
				var appInfo = channel.TeamsApp.DisplayName + " app id: " + channel.Id;
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(appInfo), cancellationToken);
			}

		}

		private async Task SendEmailToUser(WaterfallStepContext stepContext, CancellationToken cancellationToken, string emailTo, string subject, string message)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);
			var me = await client.GetMeAsync();

			try
			{
				await client.SendMailAsync(emailTo, subject, message);
			}
			catch (TeamsException e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

			await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Email to user {emailTo} was send from {me.UserPrincipalName}"), cancellationToken);
		}

		private async Task SendMessageToChanel(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string channelName, string message)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			try
			{
				await client.Teams.SendMessageToChanel(teamName, channelName, message);
			}
			catch (TeamsException e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

			await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Message was send"), cancellationToken);
		}

		private async Task RemoveTeam(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			try
			{
				await client.Teams.RemoveTeam(teamName);
			}
			catch (TeamsException e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

			await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Team with name: {teamName} was removed"), cancellationToken);
		}

		private async Task RemoveChanel(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string channelName)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			try
			{
				await client.Teams.RemoveChannelFromTeam(teamName, channelName);
			}
			catch (TeamsException e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

			await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Channel with name: {channelName} for team: {teamName} was removed"), cancellationToken);
		}

		private async Task MemberOfChanel(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string channelName)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			List<ConversationMember> users;
			try
			{
				users = await client.Teams.GetMembersOfChannelFromTeam(teamName, channelName);
			}
			catch (TeamsException e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

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

			try
			{
				await client.Teams.CreateChannelForTeam(teamName, channelName);
			}
			catch (TeamsException e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

			await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Channel with name: {channelName} for team: {teamName} was created"), cancellationToken);
		}

		private async Task CreateTeamFor(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string userEmail)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			try
			{
				await client.Teams.CreateTeamFor(teamName, userEmail);
			}
			catch (TeamsException e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

			await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Team with name: {teamName} for {userEmail} was created"), cancellationToken);
		}

		private async Task ChanelOfTeam(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			ITeamChannelsCollectionPage channels;
			try
			{
				channels = await client.Teams.GetChannelsOfTeams(teamName);
			}
			catch (TeamsException e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

			foreach (var channel in channels)
			{
				var chanelInfo = channel.DisplayName + "\nLink: " + channel.WebUrl;
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(chanelInfo), cancellationToken);
			}
		}

		private async Task MemeberOfTeam(WaterfallStepContext stepContext, CancellationToken cancellationToken, string testTeam)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			IEnumerable<TeamsUserInfo> users;
			try
			{
				users = await client.Teams.GetMembersOfTeams(testTeam);
			}
			catch (TeamsException e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

			foreach (var user in users)
			{
				var userInfo = user.DisplayName + " : " + user.Activity + " " + user.ColorEmoji;
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(userInfo), cancellationToken);
			}
		}

		private async Task ShowAllTeams(WaterfallStepContext stepContext, CancellationToken cancellationToken)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			IGraphServiceGroupsCollectionPage teams;
			try
			{
				teams = await client.Teams.GetListTeams();
			}
			catch (TeamsException e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}
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

			IEnumerable<TeamsUserInfo> users;
			try
			{
				users = await client.GetUsers();
			}
			catch (TeamsException e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

			foreach (var user in users)
			{
				var userInfo = user.DisplayName + " : " + user.Activity + " " + user.ColorEmoji;
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(userInfo), cancellationToken);
			}
		}

		private async Task ShowOneDriveContents(WaterfallStepContext stepContext, CancellationToken cancellationToken)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			IDriveItemChildrenCollectionPage driveItems;

			try
			{
				driveItems = await client.OneDrive.GetRootContents();
			}
			catch (Exception e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

			var sb = new StringBuilder();

			foreach (var driveItem in driveItems.Where(x => x.File != null || x.Folder != null).OrderBy(x => x.File != null ? 1 : 0))
			{
				var displayString = driveItem.Folder != null ? $"{driveItem.Name}{Path.DirectorySeparatorChar}" : driveItem.Name;
				sb.AppendLine(displayString);
			}

            await stepContext.Context.SendActivityAsync(MessageFactory.Text(sb.ToString()), cancellationToken);
        }

		private async Task ShowOneDriveFolderContents(WaterfallStepContext stepContext, CancellationToken cancellationToken, string folderPath)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			IDriveItemChildrenCollectionPage driveItems;

			try
			{
				driveItems = await client.OneDrive.GetFolderContents(folderPath);
			}
			catch (Exception e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

            var sb = new StringBuilder();

            foreach (var driveItem in driveItems.Where(x => x.File != null || x.Folder != null).OrderBy(x => x.File != null ? 1 : 0))
            {
                var displayString = driveItem.Folder != null ? $"{driveItem.Name}{Path.DirectorySeparatorChar}" : driveItem.Name;
                sb.AppendLine(displayString);
            }

            await stepContext.Context.SendActivityAsync(MessageFactory.Text(sb.ToString()), cancellationToken);
        }

		private async Task SearchOneDrive(WaterfallStepContext stepContext, CancellationToken cancellationToken, string searchText)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			IEnumerable<DriveItem> driveItems;

			try
			{
				driveItems = await client.OneDrive.SearchOneDrive(searchText);
			}
			catch (Exception e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

            var sb = new StringBuilder();

            foreach (var driveItem in driveItems.Where(x => x.File != null || x.Folder != null).OrderBy(x => x.File != null ? 1 : 0))
            {
                var displayString = driveItem.Folder != null ? $"{driveItem.Name}{Path.DirectorySeparatorChar}" : driveItem.Name;
                sb.AppendLine(displayString);
            }

            await stepContext.Context.SendActivityAsync(MessageFactory.Text(sb.ToString()), cancellationToken);
        }

		private async Task DeleteOneDrive(WaterfallStepContext stepContext, CancellationToken cancellationToken, string filePath)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			try
			{
				await client.OneDrive.RemoveFile(filePath);
			}
			catch (Exception e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

			var response = $"{filePath} deleted.";
			await stepContext.Context.SendActivityAsync(MessageFactory.Text(response), cancellationToken);
		}

		private async Task DownloadOneDrive(WaterfallStepContext stepContext, CancellationToken cancellationToken, string filePath)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _graphServiceClient.CreateClientFromUserBeHalf(userTokeStore.Token);

			DriveItem file;

			try
			{
				file = await client.OneDrive.GetFile(filePath);
			}
			catch (Exception e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

			var attachment = new Microsoft.Bot.Schema.Attachment
			{
				ContentUrl = file.WebUrl,
				ContentType = file.File.MimeType,
				Name = file.Name,
			};
			await stepContext.Context.SendActivityAsync(MessageFactory.Attachment(attachment), cancellationToken);
		}

		#endregion GraphMessageHandlers
	}
}
