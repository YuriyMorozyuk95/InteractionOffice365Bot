using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using AdaptiveCards;
using AdaptiveCards.Templating;
using InteractionOfficeBot.Core.Model;
using Microsoft.Bot.Schema;
using Newtonsoft.Json;

namespace InteractionOfficeBot.WebApi.Services
{
	public class CardFactory
	{
		public static async Task<AdaptiveCardTemplate> GetUserCardTemplate(CancellationToken cancellationToken)
		{
			var path = Path.Combine(Environment.CurrentDirectory, @"AdaptiveCard", "UserCard.json");
			var templateJson = await File.ReadAllTextAsync(path, cancellationToken);

			return new AdaptiveCardTemplate(templateJson);
		}

		public static Activity CreateUserActivity(AdaptiveCardTemplate cardTemplate, TeamsUserInfo teamsUserInfo)
		{
			var cardJson = cardTemplate.Expand(teamsUserInfo);

			var attachment = new Microsoft.Bot.Schema.Attachment()
			{
				ContentType = AdaptiveCard.ContentType,
				Content = JsonConvert.DeserializeObject(cardJson),
			};

			var activity = new Activity
			{
				Attachments = new List<Microsoft.Bot.Schema.Attachment>() { attachment },
				Type = ActivityTypes.Message
			};

			return activity;
		}
	}
}
