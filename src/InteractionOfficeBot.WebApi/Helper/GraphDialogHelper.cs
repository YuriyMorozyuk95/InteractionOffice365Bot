using InteractionOfficeBot.Core.Exception;
using InteractionOfficeBot.Core.Model;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder;
using Microsoft.Graph;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System;
using System.Linq;
using InteractionOfficeBot.WebApi.Services;
using InteractionOfficeBot.Core.MsGraph;
using System.Reflection;
using AdaptiveCards.Templating;
using Newtonsoft.Json;
using File = System.IO.File;
using AdaptiveCards;
using Microsoft.Bot.Schema;
using InteractionOfficeBot.Core.Extensions;

namespace InteractionOfficeBot.WebApi.Helper
{
	public interface IGraphDialogHelper
	{
		Task ShowInstalledAppForUser(WaterfallStepContext stepContext, CancellationToken cancellationToken, string email);
		Task SendEmailToUser(WaterfallStepContext stepContext, CancellationToken cancellationToken, string emailTo, string subject, string message);
		Task SendMessageToChanel(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string channelName, string message);
		Task RemoveTeam(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName);
		Task RemoveChanel(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string channelName);
		Task MemberOfChanel(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string channelName);
		Task CreateChanelForTeam(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string channelName);
		Task CreateTeamFor(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string userEmail);
		Task ChanelOfTeam(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName);
		Task MembersOfTeam(WaterfallStepContext stepContext, CancellationToken cancellationToken, string testTeam);
		Task ShowAllTeams(WaterfallStepContext stepContext, CancellationToken cancellationToken);
		Task ShowAllUsers(WaterfallStepContext stepContext, CancellationToken cancellationToken);
		Task ShowOneDriveContents(WaterfallStepContext stepContext, CancellationToken cancellationToken);
		Task ShowOneDriveFolderContents(WaterfallStepContext stepContext, CancellationToken cancellationToken, string folderPath);
		Task SearchOneDrive(WaterfallStepContext stepContext, CancellationToken cancellationToken, string searchText);
		Task DeleteOneDrive(WaterfallStepContext stepContext, CancellationToken cancellationToken, string filePath);
		Task DownloadOneDrive(WaterfallStepContext stepContext, CancellationToken cancellationToken, string filePath);
		Task GetAllTodoTasks(WaterfallStepContext stepContext, CancellationToken cancellationToken);
		Task GetTodoUpcomingTask(WaterfallStepContext stepContext, CancellationToken cancellationToken, DateTime reminderTime);
		Task CreateTodoTask(WaterfallStepContext stepContext, CancellationToken cancellationToken, string title, DateTime reminderTime);
	}
	internal class GraphDialogHelper : IGraphDialogHelper
	{
		private readonly IStateService _stateService;
		private readonly IGraphServiceClientFactory _factory;

		public GraphDialogHelper(IStateService stateService, IGraphServiceClientFactory factory)
		{
			_stateService = stateService;
			_factory = factory;
		}

		public async Task ShowInstalledAppForUser(WaterfallStepContext stepContext, CancellationToken cancellationToken, string email)
		{
			var client = await GetClient(stepContext, cancellationToken);

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

		public async Task SendEmailToUser(WaterfallStepContext stepContext, CancellationToken cancellationToken, string emailTo, string subject, string message)
		{
			var client = await GetClient(stepContext, cancellationToken);

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

		public async Task SendMessageToChanel(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string channelName, string message)
		{
			var client = await GetClient(stepContext, cancellationToken);

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

		public async Task RemoveTeam(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName)
		{
			var client = await GetClient(stepContext, cancellationToken);

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

		public async Task RemoveChanel(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string channelName)
		{
			var client = await GetClient(stepContext, cancellationToken);

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

		public async Task MemberOfChanel(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string channelName)
		{
			var client = await GetClient(stepContext, cancellationToken);

			IEnumerable<TeamsUserInfo> users;
			try
			{
				users = await client.Teams.GetMembersOfChannelFromTeam(teamName, channelName);
			}
			catch (TeamsException e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

			var template = await CardFactory.GetUserCardTemplate(cancellationToken);

			foreach (var user in users)
			{
				var activity = CardFactory.CreateUserActivity(template, user);
				await stepContext.Context.SendActivityAsync(activity, cancellationToken);
			}
		}

		public async Task CreateChanelForTeam(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string channelName)
		{
			var client = await GetClient(stepContext, cancellationToken);

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

		public async Task CreateTeamFor(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName, string userEmail)
		{
			var client = await GetClient(stepContext, cancellationToken);

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

		public async Task ChanelOfTeam(WaterfallStepContext stepContext, CancellationToken cancellationToken, string teamName)
		{
			var client = await GetClient(stepContext, cancellationToken);

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

		public async Task MembersOfTeam(WaterfallStepContext stepContext, CancellationToken cancellationToken, string testTeam)
		{
			var client = await GetClient(stepContext, cancellationToken);

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

			var template = await CardFactory.GetUserCardTemplate(cancellationToken);

			foreach (var user in users)
			{
				var activity = CardFactory.CreateUserActivity(template, user);
				await stepContext.Context.SendActivityAsync(activity, cancellationToken);

			}
		}

		public async Task ShowAllTeams(WaterfallStepContext stepContext, CancellationToken cancellationToken)
		{
			var client = await GetClient(stepContext, cancellationToken);

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

		public async Task ShowAllUsers(WaterfallStepContext stepContext, CancellationToken cancellationToken)
		{
			var client = await GetClient(stepContext, cancellationToken);

			IEnumerable<TeamsUserInfo> users;
			try
			{
				users = await client.Teams.GetUsers();
			}
			catch (TeamsException e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

			var template = await CardFactory.GetUserCardTemplate(cancellationToken);

			foreach (var user in users)
			{
				var activity = CardFactory.CreateUserActivity(template, user);
				await stepContext.Context.SendActivityAsync(activity, cancellationToken);
			}
		}

		public async Task ShowOneDriveContents(WaterfallStepContext stepContext, CancellationToken cancellationToken)
		{
			var client = await GetClient(stepContext, cancellationToken);

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

			foreach (var driveItem in driveItems.Where(x => x.IsFileOrFolder()).OrderBy(x => x.File != null ? 1 : 0))
			{
                var displayString = $"<a href=\"{driveItem.WebUrl}\">{driveItem.GetDisplayName()}</a><br/>";
                sb.Append(displayString);
			}

			await stepContext.Context.SendActivityAsync(MessageFactory.Text(sb.ToString()), cancellationToken);
		}

		public async Task ShowOneDriveFolderContents(WaterfallStepContext stepContext, CancellationToken cancellationToken, string folderPath)
		{
			var client = await GetClient(stepContext, cancellationToken);

			IDriveItemChildrenCollectionPage driveItems;

			try
			{
				driveItems = await client.OneDrive.GetFolderContents(folderPath);
			}
			catch (ServiceException e)
			{
				string message;

				if (e.StatusCode == System.Net.HttpStatusCode.NotFound)
				{
					message = "File not found";
                }
				else
				{
					message = e.Message;
				}

                await stepContext.Context.SendActivityAsync(MessageFactory.Text(message), cancellationToken);
				return;
            }
			catch (Exception e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

			var sb = new StringBuilder();

			foreach (var driveItem in driveItems.Where(x => x.IsFileOrFolder()).OrderBy(x => x.File != null ? 1 : 0))
			{
				var displayString = $"<a href=\"{driveItem.WebUrl}\">{driveItem.GetDisplayName()}</a><br/>";
				sb.Append(displayString);
			}

			await stepContext.Context.SendActivityAsync(MessageFactory.Text(sb.ToString()), cancellationToken);
		}

		public async Task SearchOneDrive(WaterfallStepContext stepContext, CancellationToken cancellationToken, string searchText)
		{
			var client = await GetClient(stepContext, cancellationToken);

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

			foreach (var driveItem in driveItems.Where(x => x.IsFileOrFolder()).OrderBy(x => x.File != null ? 1 : 0))
			{
                var displayString = $"<a href=\"{driveItem.WebUrl}\">{driveItem.GetFullPath()}</a><br/>";
                sb.Append(displayString);
			}

			await stepContext.Context.SendActivityAsync(MessageFactory.Text(sb.ToString()), cancellationToken);
		}

		public async Task DeleteOneDrive(WaterfallStepContext stepContext, CancellationToken cancellationToken, string filePath)
		{
			var client = await GetClient(stepContext, cancellationToken);

			try
			{
				await client.OneDrive.RemoveFile(filePath);
			}
            catch (ServiceException e)
            {
                string message;

                if (e.StatusCode == System.Net.HttpStatusCode.NotFound)
                {
                    message = "File not found";
                }
                else
                {
                    message = e.Message;
                }

                await stepContext.Context.SendActivityAsync(MessageFactory.Text(message), cancellationToken);
                return;
            }
            catch (Exception e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

            var response = $"{filePath} deleted.";
			await stepContext.Context.SendActivityAsync(MessageFactory.Text(response), cancellationToken);
		}

		public async Task DownloadOneDrive(WaterfallStepContext stepContext, CancellationToken cancellationToken, string filePath)
		{
			var client = await GetClient(stepContext, cancellationToken);

			DriveItem file;

			try
			{
				file = await client.OneDrive.GetFile(filePath);
			}
            catch (ServiceException e)
            {
                string message;

                if (e.StatusCode == System.Net.HttpStatusCode.NotFound)
                {
                    message = "File not found";
                }
                else
                {
                    message = e.Message;
                }

                await stepContext.Context.SendActivityAsync(MessageFactory.Text(message), cancellationToken);
                return;
            }
            catch (Exception e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

			var reply = MessageFactory.Text($"<a href=\"{file.WebUrl}\">{file.Name}</a>");

			await stepContext.Context.SendActivityAsync(reply, cancellationToken);
		}

		public async Task GetAllTodoTasks(WaterfallStepContext stepContext, CancellationToken cancellationToken)
		{
			var client = await GetClient(stepContext, cancellationToken);

			List<TodoTaskEntity> todoTask;
			try
			{
				todoTask = await client.TodoTask.GetTodoTasks();
			}
			catch (TeamsException e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}
			foreach (var task in todoTask)
			{
				var taskInfo = task.Title + " Status: " + task.Status + " ReminderDateTime: " + task.ReminderDateTime;
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(taskInfo), cancellationToken);
			}
		}

		public async Task GetTodoUpcomingTask(WaterfallStepContext stepContext, CancellationToken cancellationToken, DateTime reminderTime)
		{
			var client = await GetClient(stepContext, cancellationToken);

			List<TodoTaskEntity> upcomingTodoTask;
			try
			{
				upcomingTodoTask = await client.TodoTask.GetUpcomingTodoTasks(reminderTime);
			}
			catch (TeamsException e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}
			foreach (var task in upcomingTodoTask)
			{
				var taskInfo = task.Title + " Status: " + task.Status + " ReminderDateTime: " + task.ReminderDateTime;
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(taskInfo), cancellationToken);
			}
		}

		public async Task CreateTodoTask(WaterfallStepContext stepContext, CancellationToken cancellationToken, string title, DateTime reminderTime)
		{
			var client = await GetClient(stepContext, cancellationToken);

			try
			{
				await client.TodoTask.CreateTodoTask(title, reminderTime);
			}
			catch (TeamsException e)
			{
				await stepContext.Context.SendActivityAsync(MessageFactory.Text(e.Message), cancellationToken);
				return;
			}

			await stepContext.Context.SendActivityAsync(MessageFactory.Text($"ToDo Task with title: {title} was created"), cancellationToken);
		}

		private async Task<IobGraphClient> GetClient(WaterfallStepContext stepContext, CancellationToken cancellationToken)
		{
			var userTokeStore = await _stateService.UserTokeStoreAccessor.GetAsync(stepContext.Context, () => new UserTokeStore(), cancellationToken);
			var client = _factory.CreateClientFromUserBeHalf(userTokeStore.Token);
			return client;
		}

	}
}
