using Microsoft.Bot.Builder.AI.Luis;
using Microsoft.Extensions.Configuration;

using System.Threading.Tasks;
using InteractionOfficeBot.WebApi.Model;
using Microsoft.Bot.Builder;
using System.Threading;

namespace InteractionOfficeBot.WebApi.Services
{
	public interface ILuisService
	{
		Task<LuisRoot> Recognize(ITurnContext turnContext, CancellationToken cancellationToken = default);
	}

	public class LuisService : ILuisService
	{
		private readonly LuisRecognizer _dispatch;

		public LuisService(IConfiguration configuration)
		{
			// Read the setting for cognitive services (LUIS, QnA) from the appsettings.json
			var luisApplication = new LuisApplication(
				configuration["LuisAppId"],
				configuration["LuisAPIKey"],
				configuration["LuisAPIWeb"]);

			var recognizerOptions = new LuisRecognizerOptionsV3(luisApplication)
			{
				PredictionOptions = new Microsoft.Bot.Builder.AI.LuisV3.LuisPredictionOptions
				{
					IncludeAllIntents = true,
					IncludeInstanceData = true,
				}
			};

			_dispatch = new LuisRecognizer(recognizerOptions);
		}

		public Task<LuisRoot> Recognize(ITurnContext turnContext, CancellationToken cancellationToken = default)
		{
			return _dispatch.RecognizeAsync<LuisRoot>(turnContext, cancellationToken);
		}
	}
}
