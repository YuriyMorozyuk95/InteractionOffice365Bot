using Microsoft.Graph;

using System.Threading.Tasks;
using InteractionOfficeBot.Core.Model;

namespace InteractionOfficeBot.Core.MsGraph
{
	// This class is a wrapper for the Microsoft Graph API
	public class IobGraphClient
    {
	    private readonly GraphServiceClient _graphClient;

	    private TeamsRepository? _teamsGroupRepository;

        private OneDriveRepository? _oneDriveRepository;

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

        // Sends an email on the users behalf using the Microsoft Graph API
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

        // Get information about the user.
        public async Task<IEnumerable<TeamsUserInfo>> GetUsers()
        {
	        var users = await _graphClient.Users
		        .Request()
		        .GetAsync();;

	        var result = await _graphClient.Communications.GetPresencesByUserId(users.Select(x => x.Id)).Request().PostAsync();

	        var teamsUserInfo = result.Select(x => new TeamsUserInfo
	        {
		        DisplayName = users.FirstOrDefault(m => m.Id == x.Id)?.DisplayName,
		        Activity = x.Activity,
	        }).OrderBy(x => x.Activity);

	        return teamsUserInfo;
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

        // // Gets the user's photo
        // public async Task<PhotoResponse> GetPhotoAsync()
        // {
        //     HttpClient client = new HttpClient();
        //     client.DefaultRequestHeaders.Add("Authorization", "Bearer " + _token);
        //     client.DefaultRequestHeaders.Add("Accept", "application/json");

        //     using (var response = await client.GetAsync("https://graph.microsoft.com/v1.0/me/photo/$value"))
        //     {
        //         if (!response.IsSuccessStatusCode)
        //         {
        //             throw new HttpRequestException($"Graph returned an invalid success code: {response.StatusCode}");
        //         }

        //         var stream = await response.Content.ReadAsStreamAsync();
        //         var bytes = new byte[stream.Length];
        //         stream.Read(bytes, 0, (int)stream.Length);

        //         var photoResponse = new PhotoResponse
        //         {
        //             Bytes = bytes,
        //             ContentType = response.Content.Headers.ContentType?.ToString(),
        //         };

        //         if (photoResponse != null)
        //         {
        //             photoResponse.Base64String = $"data:{photoResponse.ContentType};base64," +
        //                                          Convert.ToBase64String(photoResponse.Bytes);
        //         }

        //         return photoResponse;
        //     }
        // }
    }
}
