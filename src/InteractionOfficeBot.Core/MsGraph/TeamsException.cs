namespace InteractionOfficeBot.Core.MsGraph;

public class TeamsException : ApplicationException
{
	public TeamsException()
	{
	}

	public TeamsException(string message) : base(message)
	{
	}

	public TeamsException(string message, Exception inner) : base(message, inner)
	{
	}
}