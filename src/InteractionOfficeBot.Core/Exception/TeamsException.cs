namespace InteractionOfficeBot.Core.Exception;

public class TeamsException : ApplicationException
{
    public TeamsException()
    {
    }

    public TeamsException(string message) : base(message)
    {
    }

    public TeamsException(string message, System.Exception inner) : base(message, inner)
    {
    }
}