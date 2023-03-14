namespace InteractionOfficeBot.Core.Model;

public class TeamsUserInfo
{
	public string? DisplayName { get; set; }
	public string? Activity { get; set; }

	//TODO card color
	public string Color
	{
		get
		{
			switch (Activity)
			{
				case "Available":
					return "Green";
				case "Busy":
				case "DoNotDisturb":
					return "Red";
				case "BeRightBack":
				case "Away":
					return "Yellow";
				default :
					return "Grey";
			}
		}
	}
}