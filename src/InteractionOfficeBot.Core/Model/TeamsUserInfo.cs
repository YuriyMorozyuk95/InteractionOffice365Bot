namespace InteractionOfficeBot.Core.Model;

public class TeamsUserInfo
{
	public string? DisplayName { get; set; }
	public string? Activity { get; set; }
	public string? ImageUrl { get; set; }

	public string ColorEmoji
	{
		get
		{
			switch (Activity)
			{
				case "Available":
					return "🟢";
				case "Busy":
				case "DoNotDisturb":
				case "InACall":
				case "InAConferenceCall":
				case "InAMeeting":
				case "Presenting":
				case "UrgentInterruptionsOnly":
					return "🔴";
				case "BeRightBack":
				case "Away":
					return "🟡";
				default :
					return "🔘";
			}
		}
	}
}