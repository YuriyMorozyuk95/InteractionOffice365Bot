using InteractionOfficeBot.Core.Exception;
using InteractionOfficeBot.Core.Model;
using InteractionOfficeBot.WebApi.Services;

using System;
using System.Linq;

namespace InteractionOfficeBot.WebApi.Helper
{
	internal class LuisEntityHelper
	{
		public static string GetFileFromEntity(LuisRoot recognizeResult)
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
		public static string GetFolderPathFromEntity(LuisRoot recognizeResult)
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
		public static string GetTeamFromEntity(LuisRoot recognizeResult)
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
		public static string GetChannelFromEntity(LuisRoot recognizeResult)
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
		public static string GetUserFromEntity(LuisRoot recognizeResult)
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
		public static string GetMessageFromEntity(LuisRoot recognizeResult)
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
		public static string GetEmailSubjectFromEntity(LuisRoot recognizeResult)
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
		public static string GetTaskTitleFromEntity(LuisRoot recognizeResult)
		{
			var title = recognizeResult.Entities
				?.Title
				?.FirstOrDefault()
				?.Value
				?.FirstOrDefault();

			if (title == null)
			{
				throw new TeamsException("Can't recognize task");
			}

			return title;
		}
		public static DateTime? GetTaskReminderTimeFromEntity(LuisRoot recognizeResult)
		{
			var reminderTime = recognizeResult.Entities
				?.ReminderTime
				?.FirstOrDefault()
				?.Value
				?.FirstOrDefault();

			if (reminderTime == null)
			{
				throw new TeamsException("Can't recognize reminder time");
			}



			var dt = AiRecognizer.RecognizeDateTime(reminderTime, out _);
			return dt;
		}
	}
}
