using Microsoft.Graph;

namespace InteractionOfficeBot.Core.MsGraph
{
	// This class is a wrapper for the Microsoft Graph API
	public class IobGraphClient
    {
	    private readonly GraphServiceClient _graphClient;

	    private TeamsGroupRepository? _teamsGroupRepository;

        public IobGraphClient(GraphServiceClient graphServiceClient)
        {
	        _graphClient = graphServiceClient;
        }

        public TeamsGroupRepository TeamsGroup
        {
	        get { return _teamsGroupRepository ??= new TeamsGroupRepository(_graphClient); }
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
                new Recipient
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
        public Task<IGraphServiceUsersCollectionPage> GetUsers()
        {
	        var graphRequest = _graphClient.Users
		        .Request()
		        .Select(u => new { u.DisplayName, u.Mail });

	        return graphRequest.GetAsync();
        }

        public async Task<User> GetMyUsers()
        {
            //TODO add const
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

	public class TeamsGroupRepository
	{
		private readonly GraphServiceClient _graphServiceClient;

		public TeamsGroupRepository(GraphServiceClient graphServiceClient)
		{
			_graphServiceClient = graphServiceClient;
		}



		public Task<IGraphServiceGroupsCollectionPage> List()
		{
			var graphRequest = _graphServiceClient.Groups
				.Request()
				.Filter("resourceProvisioningOptions/Any(x:x eq 'Team')")
				.Select(x => new { x.DisplayName, x.Id });

			return graphRequest.GetAsync();
		}

		public async Task Create(User user = null)
		{
			user ??= await _graphServiceClient.Me.Request().GetAsync();

			var groupRequestBody = new Group
			{
				Description = "Self help community for library 2",
				DisplayName = "Library Assist",
				GroupTypes = new List<string>
				{
					"Unified",
				},
				MailEnabled = true,
				MailNickname = "library",
				SecurityEnabled = false,
			};

			//var groupRequest = await _graphServiceClient.Groups.Request().AddAsync(groupRequestBody);

			var requestBody = new Team
			{
                //Group = groupRequestBody,
				DisplayName = "Architecture Team 3",
				Description = "The team for those in architecture design.",
				AdditionalData = new Dictionary<string, object>
				{
					{
						"template@odata.bind" , "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
					},
				},
				Members = new TeamMembersCollectionPage()  
				{  
					new AadUserConversationMember  
					{  
						Roles = new List<String>()  
						{  
							"owner"  
						},  
						AdditionalData = new Dictionary<string, object>()  
						{  
							{"user@odata.bind", $"https://graph.microsoft.com/v1.0/users('{user.Id}')"}  
						}  
					}  
				},  
			};

			var result = await _graphServiceClient.Teams
				.Request()
				.AddAsync(requestBody);
		}
    }
}
