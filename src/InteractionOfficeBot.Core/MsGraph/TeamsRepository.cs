﻿using System.Text.Json.Serialization;

using InteractionOfficeBot.Core.Exception;
using InteractionOfficeBot.Core.Model;

using Microsoft.Graph;

using File = System.IO.File;

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


	// Get information about the user.
	public async Task<IEnumerable<TeamsUserInfo>> GetUsers()
	{
		var users = await _graphServiceClient.Users
			.Request()
			.GetAsync(); ;

		var result = await _graphServiceClient.Communications.GetPresencesByUserId(users.Select(x => x.Id)).Request().PostAsync();

		var list = new List<TeamsUserInfo>();
		foreach (var user in users)
		{
			var actualUrl = await GetUserPhotoUrl(user.Id);

			list.Add(new TeamsUserInfo
			{
				DisplayName = user.DisplayName,
				Activity = result.FirstOrDefault(m => user.Id == m.Id)?.Activity,
				ImageUrl = actualUrl,
			});
		}

		return list.OrderBy(x => x.Activity);
	}

	public async Task<IEnumerable<TeamsUserInfo>> GetMembersOfTeams(string teamName)
	{
		var group = await GetTeam(teamName);

		var members = await _graphServiceClient
			.Teams[group.Id]
			.Members
			.Request()
			.GetAsync();

		var result = await _graphServiceClient.Communications
			.GetPresencesByUserId(members.Select(x => ((AadUserConversationMember)x).UserId))
			.Request()
			.PostAsync();

		var list = new List<TeamsUserInfo>();
		foreach (var member in members)
		{
			var adMember = member as AadUserConversationMember;
			var actualUrl = await GetUserPhotoUrl(adMember?.UserId);

			list.Add(new TeamsUserInfo
			{
				DisplayName = member.DisplayName,
				Activity = result.FirstOrDefault(m => adMember?.UserId == m.Id)?.Activity,
				ImageUrl = actualUrl,
			});
		}

		return list.OrderBy(x => x.Activity);
	}

	public async Task<IEnumerable<TeamsUserInfo>> GetMembersOfChannelFromTeam(string teamName, string chanelName)
	{
		var group = await GetTeam(teamName);

		var channel = await GetChannel(group.Id, chanelName);

		var members = await _graphServiceClient
			.Teams[group.Id]
			.Channels[channel.Id]
			.Members
			.Request()
			.GetAsync();

		var result = await _graphServiceClient.Communications.GetPresencesByUserId(members.Select(x => ((AadUserConversationMember)x).UserId))
			.Request()
			.PostAsync();

		var list = new List<TeamsUserInfo>();
		foreach (var member in members)
		{
			var adMember = member as AadUserConversationMember;
			var actualUrl = await GetUserPhotoUrl(adMember?.UserId);

			list.Add(new TeamsUserInfo
			{
				DisplayName = member.DisplayName,
				Activity = result.FirstOrDefault(m => adMember?.UserId == m.Id)?.Activity,
				ImageUrl = actualUrl,
			});
		}

		return list.OrderBy(x => x.Activity);
	}

	public async Task<ITeamChannelsCollectionPage> GetChannelsOfTeams(string teamName)
	{
		var group = await GetTeam(teamName);

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



	public async Task RemoveChannelFromTeam(string teamName, string chanelName)
	{
		var groups = await GetTeams(teamName);

		var group = groups.FirstOrDefault();

		if (group == null)
		{
			throw new TeamsException($"There is no team with name {teamName}");
		}

		var channels = await _graphServiceClient
			.Me
			.Todo
			.Lists
			.Request()
			.Filter($"displayName eq '{chanelName}'")
			.GetAsync();

		var channel = channels.FirstOrDefault();

		if (channel == null)
		{
			throw new TeamsException($"There is no channel with name {teamName}");
		}

		await _graphServiceClient
			.Teams[group.Id]
			.Channels[channel.Id]
			.Request()
			.DeleteAsync();
	}

	public async Task RemoveTeam(string teamName)
	{
		var groups = await GetTeams(teamName);

		var group = groups.FirstOrDefault();

		if (group == null)
		{
			throw new TeamsException($"There is no team with name {teamName}");
		}

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

	public async Task<IUserTeamworkInstalledAppsCollectionPage> GetInstalledAppForUser(string userEmail)
	{
		return await _graphServiceClient.Users[userEmail]
			.Teamwork
			.InstalledApps
			.Request()
			.Expand(item => item.TeamsApp)
			.Select(x => new { x.Id, x.TeamsApp })
			.GetAsync();
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

	private async Task<string> GetUserPhotoUrl(string userId)
	{
		byte[] imageBytes;
		try
		{
			var photoStream = await _graphServiceClient.Users[userId].Photo.Content.Request().GetAsync();
			imageBytes = new byte[photoStream.Length];
			await photoStream.ReadAsync(imageBytes, 0, imageBytes.Length);
		}
		catch
		{
			var path = Path.Combine(Environment.CurrentDirectory, @"Img", "FunnyAvatar.png");
			imageBytes = await File.ReadAllBytesAsync(path);
		}

		string actualUrl = "data:image/gif;base64," + Convert.ToBase64String(imageBytes);
		return actualUrl;
	}
}