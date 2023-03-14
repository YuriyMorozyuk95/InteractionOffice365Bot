//using Newtonsoft.Json.Linq;
//using System.Collections.Generic;
//using System;
//using Microsoft.Bot.Schema;

//namespace InteractionOfficeBot.WebApi
//{
//	public class CardFactory
//	{
//		private const string AdaptiveCardSchemaVersion = "1.3";
//		public const string DealCard = "DealCard";

//		public static Attachment CreateDealCard(DealCreationDataModel model)
//		{
//			var dealCardId =  $"{DealCard}_{Guid.NewGuid()}";
//            var locationEditorId =  $"{nameof(DealCreationDataModelEditor.LocationEditor)}_{Guid.NewGuid()}";

//            var periodFromDateEditorId =  $"{nameof(DealCreationDataModelEditor.PeriodFromDateEditor)}_{Guid.NewGuid()}";
//            var periodFromTimeEditorId =  $"{nameof(DealCreationDataModelEditor.PeriodFromTimeEditor)}_{Guid.NewGuid()}";

//            var periodToDateEditorId =  $"{nameof(DealCreationDataModelEditor.PeriodToDateEditor)}_{Guid.NewGuid()}";
//            var periodToTimeEditorId =  $"{nameof(DealCreationDataModelEditor.PeriodToTimeEditor)}_{Guid.NewGuid()}";

//            var quantityDateEditorId =  $"{nameof(DealCreationDataModelEditor.QuantityDateEditor)}_{Guid.NewGuid()}";
//            var priceDateEditorId =  $"{nameof(DealCreationDataModelEditor.PriceDateEditor)}_{Guid.NewGuid()}";
//            var userEmailEditorId =  $"{nameof(DealCreationDataModelEditor.UserEmailEditor)}_{Guid.NewGuid()}";

//            var cancelId = $"{nameof(Models.Action.Cancel)}_{Guid.NewGuid()}";
//            var confirmId = $"{nameof(Models.Action.Confirm)}_{Guid.NewGuid()}";


//			var adaptiveCard = new AdaptiveCard(new AdaptiveSchemaVersion(AdaptiveCardSchemaVersion))
//			{
//				Id = dealCardId,
//				Body = new List<AdaptiveElement>
//				{
//					new AdaptiveTextInput
//					{
//						Label = "Location:",
//						Id = locationEditorId,
//						Value = model.Location
//					},
//					new AdaptiveContainer
//					{
//						Items = new List<AdaptiveElement>
//						{
//							new AdaptiveTextBlock
//							{
//								Text = "Period From:"
//							},
//							new AdaptiveDateInput
//							{
//								Id = periodFromDateEditorId,
//								Value = model.Period.From.ToString("yyyy-MM-dd"),
//							},
//							new AdaptiveTimeInput
//							{
//								Id = periodFromTimeEditorId,
//								Value = model.Period.From.ToString("HH:mm"),
//							},
//						},
//					},
//					new AdaptiveContainer
//					{
//						Items = new List<AdaptiveElement>
//						{
//							new AdaptiveTextBlock
//							{
//								Text = "Period To:"
//							},
//							new AdaptiveDateInput
//							{
//								Id = periodToDateEditorId,
//								Value = model.Period.To.ToString("yyyy-MM-dd")
//							},
//							new AdaptiveTimeInput
//							{
//								Id = periodToTimeEditorId,
//								Value = model.Period.To.ToString("HH:mm"),
//							},
//						},
//					},
//					new AdaptiveNumberInput
//					{
//						Label = "Quantity:",
//						Id = quantityDateEditorId,
//						Value = (double)model.Quantity,
//					},
//					new AdaptiveNumberInput
//					{
//						Label = "Price:",
//						Id = priceDateEditorId,
//						Value = (double)model.Price,
//					},
//					new AdaptiveTextInput
//					{
//						Label = $"Email:",
//						Id = userEmailEditorId,
//						Value = model.UserEmail
//					},
//				},
//				Actions = new List<AdaptiveAction>
//				{
//					new AdaptiveSubmitAction()
//					{
//						Id = cancelId,
//						Type = AdaptiveSubmitAction.TypeName,
//						Title = Cancel,
//						Data = new JObject { { nameof(Models.Action), nameof(Models.Action.Cancel) } }
//					},
//					new AdaptiveSubmitAction()
//					{
//						Id = confirmId,
//						Type = AdaptiveSubmitAction.TypeName,
//						Title = Confirm,
//						Data = new JObject { { nameof(Models.Action), nameof(Models.Action.Confirm) } },
//					},
//				}
//			};

//            return new Attachment()
//            {
//                ContentType = AdaptiveCard.ContentType,
//                Content = adaptiveCard,
//            };
//		}
//	}
//}
