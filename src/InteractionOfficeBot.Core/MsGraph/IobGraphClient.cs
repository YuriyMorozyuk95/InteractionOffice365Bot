using Microsoft.Graph;

using System.Threading.Tasks;
using InteractionOfficeBot.Core.Model;
using File = System.IO.File;

namespace InteractionOfficeBot.Core.MsGraph
{
	// This class is a wrapper for the Microsoft Graph API
	public class IobGraphClient
    {
	    private readonly GraphServiceClient _graphClient;

	    private TeamsRepository? _teamsGroupRepository;

        private OneDriveRepository? _oneDriveRepository;

        private TodoTaskRepository? _todoGroupRepository;

        public IobGraphClient(GraphServiceClient graphServiceClient)
        {
	        _graphClient = graphServiceClient;
        }

        public GraphServiceClient Client => _graphClient;

        public TeamsRepository Teams
        {
	        get { return _teamsGroupRepository ??= new TeamsRepository(_graphClient); }
        }

        public OneDriveRepository OneDrive
        {
            get => _oneDriveRepository ??= new OneDriveRepository(_graphClient);
        }

        public TodoTaskRepository TodoTask
        {
            get { return _todoGroupRepository ??= new TodoTaskRepository(_graphClient); }
        }

        public async Task SendMailAsync(string toAddress, string subject, string content)
        {
            if (string.IsNullOrWhiteSpace(toAddress))
            {
                throw new ArgumentNullException(nameof(toAddress));
            }

            if (string.IsNullOrWhiteSpace(subject))
            {
                throw new ArgumentNullException(nameof(subject));
            }

            if (string.IsNullOrWhiteSpace(content))
            {
                throw new ArgumentNullException(nameof(content));
            }

            var recipients = new List<Recipient>
            {
                new()
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = toAddress,
                    },
                },
            };

            // Create the message.
            var email = new Message
            {
                Body = new ItemBody
                {
                    Content = content,
                    ContentType = BodyType.Text,
                },
                Subject = subject,
                ToRecipients = recipients,
            };

            // Send the message.
            await _graphClient.Me.SendMail(email, true).Request().PostAsync();
        }

        // Gets mail for the user using the Microsoft Graph API
        public async Task<Message[]> GetRecentMailAsync()
        {
            var messages = await _graphClient.Me.MailFolders.Inbox.Messages.Request().GetAsync();
            return messages.Take(5).ToArray();
        }

        // Get information about the user.
        public async Task<User> GetMeAsync()
        {
            var me = await _graphClient.Me.Request().GetAsync();
            return me;
        }

        public async Task<User> GetMyUsers()
        {
	        return await _graphClient.Users["yurii.moroziuk.iob@8bpskq.onmicrosoft.com"].Request().GetAsync();
        }

        // gets information about the user's manager.
        public async Task<User?> GetManagerAsync()
        {
            var manager = await _graphClient.Me.Manager.Request().GetAsync() as User;
            return manager;
        }
    }
}
