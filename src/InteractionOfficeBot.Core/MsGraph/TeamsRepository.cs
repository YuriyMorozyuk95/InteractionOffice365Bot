using InteractionOfficeBot.Core.Exception;
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
		var group = await GetTeam(teamName);

		//TODO validate Any
		var members = await _graphServiceClient
			.Teams[group.Id]
			.Members
			.Request()
			.GetAsync();

		return members.ToList();
	}

	public async Task<ITeamChannelsCollectionPage> GetChannelsOfTeams(string teamName)
	{
		var group = await GetTeam(teamName);

		//TODO validate Any
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
			throw new TeamsException($"user with email {userEmail} don't exist");
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

		await TeamValidateAny(teamName);

		await _graphServiceClient.Teams
			.Request()
			.AddAsync(requestBody);
	}

	public async Task CreateChannelForTeam(string teamName, string chanelName)
	{
		var group = await GetTeam(teamName);

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
		var group = await GetTeam(teamName);

		var channel = await GetChannel(group.Id, chanelName);

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
		var group = await GetTeam(teamName);

		var channel = await GetChannel(group.Id, chanelName);

		await _graphServiceClient
			.Teams[group.Id]
			.Channels[channel.Id]
			.Request()
			.DeleteAsync();
	}

	public async Task RemoveTeam(string teamName)
	{
		var group = await GetTeam(teamName);

		await _graphServiceClient
			.Groups[group.Id]
			.Request()
			.DeleteAsync();
	}

	public async Task SendMessageToChanel(string teamName, string chanelName, string message)
	{
		var group = await GetTeam(teamName);

		var channel = await GetChannel(group.Id, chanelName);

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

	private Channel ValidateAndGetChannel(string channelName, ITeamChannelsCollectionPage channels)
	{
		Channel? channel;
		try
		{
			channel = channels.SingleOrDefault();

			if (channel == null)
			{
				throw new TeamsException($"There no team with name {channelName}");
			}
		}
		catch (InvalidOperationException e)
		{
			throw new TeamsException($"There is more than one team with name {channelName}", e);
		}

		return channel;
	}

	private Group ValidateAndGetGroup(string teamName, IGraphServiceGroupsCollectionPage groups)
	{
		Group? group;
		try
		{
			group = groups.SingleOrDefault();

			if (group == null)
			{
				throw new TeamsException($"There no team with name {teamName}");
			}
		}
		catch (InvalidOperationException e)
		{
			throw new TeamsException($"There is more than one team with name {teamName}", e);
		}

		return group;
	}

	private async Task<Channel> GetChannel(string groupId, string channelName)
	{
		var channels = await _graphServiceClient
			.Teams[groupId]
			.Channels
			.Request()
			.Filter($"displayName eq '{channelName}'")
			.GetAsync();

		return ValidateAndGetChannel(channelName, channels);
	}

	private async Task<Group> GetTeam(string teamName)
	{
		var groups = await GetTeams(teamName);

		return ValidateAndGetGroup(teamName, groups);
	}

	private async Task<IGraphServiceGroupsCollectionPage> GetTeams(string teamName)
	{
		var groups = await _graphServiceClient.Groups
			.Request()
			.Filter($"resourceProvisioningOptions/Any(x:x eq 'Team') and displayName eq '{teamName}'")
			.GetAsync();
		return groups;
	}

	private async Task TeamValidateAny(string teamName)
	{
		var groups = await GetTeams(teamName);

		if (groups == null || groups?.Count == 0)
		{
			throw new TeamsException($"team with name {teamName} already exist");
		}
	}
}