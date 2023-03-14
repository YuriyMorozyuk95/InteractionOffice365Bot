namespace InteractionOfficeBot.Core.Model;

public class TeamsUserInfo
{
	public string? DisplayName { get; set; }
	public string? Activity { get; set; }

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