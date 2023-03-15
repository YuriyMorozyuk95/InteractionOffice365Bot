using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder;
using System;
using Microsoft.Extensions.Configuration;

namespace InteractionOfficeBot.WebApi.Services
{
	public interface IStateService
	{
		ConversationState ConversationState { get; }
		UserState UserState { get; }

		IStatePropertyAccessor<UserTokeStore> UserTokeStoreAccessor { get; }
		IStatePropertyAccessor<DialogState> DialogStateAccessor { get; }
		IStatePropertyAccessor<DateTime> LastAccessedTimeAccessor {get;}
	}
	internal class StateService : IStateService
	{
		#region Variables
		// State Variables
		public ConversationState ConversationState { get; }
		public UserState UserState { get; }
		// IDs
		public static string DialogStateId { get; } = $"{nameof(StateService)}.DialogState";
		public static string UserTokeStoreId { get; } = $"{nameof(StateService)}.UserTokeStore";
		public static string LastAccessedTimeId { get; } = $"{nameof(StateService)}.LastAccessedTime";

		// Accessors
		public IStatePropertyAccessor<UserTokeStore> UserTokeStoreAccessor { get; set; }
		public IStatePropertyAccessor<DialogState> DialogStateAccessor { get; set; }
		public IStatePropertyAccessor<DateTime> LastAccessedTimeAccessor { get; set; }

		#endregion
		public StateService(UserState userState, ConversationState conversationState)
		{
			ConversationState = conversationState ?? throw new ArgumentNullException(nameof(conversationState));
			UserState = userState ?? throw new ArgumentNullException(nameof(userState));

			InitializeAccessors();
		}

		public int ExpireAfterMinutes { get; set; }

		public void InitializeAccessors()
		{
			// Initialize Conversation State Accessors
			DialogStateAccessor = ConversationState.CreateProperty<DialogState>(DialogStateId);
			// Initialize User State
			UserTokeStoreAccessor = UserState.CreateProperty<UserTokeStore>(UserTokeStoreId);
			LastAccessedTimeAccessor = ConversationState.CreateProperty<DateTime>(nameof(LastAccessedTimeAccessor));
		}

	}

	//TODO to model
	public class UserTokeStore
	{
		public string Token { get; set; }
	}
}
