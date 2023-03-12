using Microsoft.Graph;

namespace InteractionOfficeBot.Core.MsGraph;

public class TeamsRepository
{
	private readonly GraphServiceClient _graphServiceClient;

	public TeamsRepository(GraphServiceClient graphServiceClient)
	{
		_graphServiceClient = graphServiceClient;
	}

	public Task<IGraphServiceGroupsCollectionPage> GetListTeams()
	{
		var graphRequest = _graphServiceClient.Groups
			.Request()
			.Filter("resourceProvisioningOptions/Any(x:x eq 'Team')");

		return graphRequest.GetAsync();
	}

	public async Task<List<ConversationMember>> GetMembersOfTeams(string teamName)
	{
		var groups = await _graphServiceClient.Groups
			.Request()
			.Filter($"resourceProvisioningOptions/Any(x:x eq 'Team') and displayName eq '{teamName}'")
			.Top(1)
			.GetAsync();

		var group = groups.Single();

		var members = await _graphServiceClient
			.Teams[group.Id]
			.Members
			.Request()
			.GetAsync();


		return members.ToList();
	}

	public async Task<ITeamChannelsCollectionPage> GetChannelsOfTeams(string teamName)
	{
		var groups = await _graphServiceClient.Groups
			.Request()
			.Filter($"resourceProvisioningOptions/Any(x:x eq 'Team') and displayName eq '{teamName}'")
			.Top(1)
			.GetAsync();

		var group = groups.Single();

		var channels = await _graphServiceClient
			.Teams[group.Id]
			.Channels
			.Request()
			.Select(x => new { x.DisplayName, x.WebUrl, })
			.GetAsync();

		return channels;
	}

	public async Task CreateTeamFor(string teamName, string userEmail)
	{
		var user = await _graphServiceClient.Users[userEmail].Request().GetAsync();

		if (user == null)
		{
			throw new TeamsException("user don't exist");
		}

		var requestBody = new Team
		{
			DisplayName = teamName,
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
					Roles = new List<string>()  
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

	public async Task CreateChannelForTeam(string teamName, string chanelName)
	{
		var groups = await _graphServiceClient.Groups
			.Request()
			.Filter($"resourceProvisioningOptions/Any(x:x eq 'Team') and displayName eq '{teamName}'")
			.Top(1)
			.GetAsync();

		var group = groups.Single();

		var requestBody = new Channel
		{
			DisplayName = chanelName,
			MembershipType = ChannelMembershipType.Standard,
		};

		await _graphServiceClient
			.Teams[group.Id]
			.Channels
			.Request()
			.AddAsync(requestBody);
	}

	public async Task<List<ConversationMember>> GetMembersOfChannelFromTeam(string teamName, string chanelName)
	{
		var groups = await _graphServiceClient.Groups
			.Request()
			.Filter($"resourceProvisioningOptions/Any(x:x eq 'Team') and displayName eq '{teamName}'")
			.Top(1)
			.GetAsync();

		var group = groups.Single();

		var channels = await _graphServiceClient
			.Teams[group.Id]
			.Channels
			.Request()
			.Filter($"displayName eq '{chanelName}'")
			.GetAsync();

		var channel = channels.First();

		var members = await _graphServiceClient
			.Teams[group.Id]
			.Channels[channel.Id]
			.Members
			.Request()
			.GetAsync();

		return members.ToList();
	}

	public async Task RemoveChannelFromTeam(string teamName, string chanelName)
	{
		var groups = await _graphServiceClient.Groups
			.Request()
			.Filter($"resourceProvisioningOptions/Any(x:x eq 'Team') and displayName eq '{teamName}'")
			.Top(1)
			.GetAsync();

		var group = groups.Single();

		var channels = await _graphServiceClient
			.Teams[group.Id]
			.Channels
			.Request()
			.Filter($"displayName eq '{chanelName}'")
			.GetAsync();

		var channel = channels.First();

		await _graphServiceClient
			.Teams[group.Id]
			.Channels[channel.Id]
			.Request()
			.DeleteAsync();
	}

	public async Task RemoveTeam(string teamName)
	{
		var groups = await _graphServiceClient.Groups
			.Request()
			.Filter($"resourceProvisioningOptions/Any(x:x eq 'Team') and displayName eq '{teamName}'")
			.Top(1)
			.GetAsync();

		var group = groups.Single();

		await _graphServiceClient
			.Groups[group.Id]
			.Request()
			.DeleteAsync();
	}

	public async Task SendMessageToChanel(string teamName, string chanelName, string message)
	{
		var groups = await _graphServiceClient.Groups
			.Request()
			.Filter($"resourceProvisioningOptions/Any(x:x eq 'Team') and displayName eq '{teamName}'")
			.Top(1)
			.GetAsync();

		var group = groups.Single();

		var channels = await _graphServiceClient
			.Teams[group.Id]
			.Channels
			.Request()
			.Filter($"displayName eq '{chanelName}'")
			.GetAsync();

		var channel = channels.First();

		var requestBody = new ChatMessage
		{
			Body = new ItemBody
			{
				Content = message,
			},
		};

		await _graphServiceClient
			.Teams[group.Id]
			.Channels[channel.Id]
			.Messages
			.Request()
			.AddAsync(requestBody);
	}
}