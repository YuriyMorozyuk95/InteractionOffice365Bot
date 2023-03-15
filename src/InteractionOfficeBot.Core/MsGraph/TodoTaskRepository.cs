using InteractionOfficeBot.Core.Model;
using Microsoft.Graph;

namespace InteractionOfficeBot.Core.MsGraph;

public class TodoTaskRepository
{
    private readonly GraphServiceClient _graphServiceClient;

    public TodoTaskRepository(GraphServiceClient graphServiceClient)
    {
        _graphServiceClient = graphServiceClient;
    }

    public async Task<List<TodoTaskEntity>> GetTodoTasks()
    {
        var listId = await GetListId();
        var result = await _graphServiceClient.Me.Todo.Lists[listId].Tasks.Request().GetAsync();

        var todoTask = result.Select(x => new TodoTaskEntity
        {
            Title = x.Title,
            ReminderDateTime = x.ReminderDateTime,
            Status = x.Status,
        }).ToList();

        return todoTask;
    }

    public async Task<List<TodoTaskEntity>> GetUpcomingTodoTasks(DateTime? reminderTime)
    {
	    string? reminderTimeFrom;
	    string? reminderTimeTo;

	    if (reminderTime != null)
	    {
		    var universalDateTime = reminderTime.Value.ToUniversalTime();

		    var from = universalDateTime.Date;
		    var to = universalDateTime.Date.AddDays(1);

		    //TODO try "u"
		    string fromat = "yyyy-MM-dd'T'HH:mm:ss'Z'";

		    var reminderTimeStart = from.ToString(fromat);
            var reminderTimeTo = to.ToString(fromat);

	    }

        var listId = await GetListId();
        var result = await _graphServiceClient
	        .Me
	        .Todo
	        .Lists[listId]
	        .Tasks
	        .Request()
	        .Filter("reminderDateTime/dateTime gt '2022-07-15T01:30:00Z'")
	        .GetAsync();

        var upcomingTasks = result
            .Select(x => new TodoTaskEntity
            {
                Title = x.Title,
                ReminderDateTime = x.ReminderDateTime,
                Status = x.Status,
            }).ToList();

        return upcomingTasks;
    }

    public async Task CreateTodoTask(string title, DateTime? reminderTime)
    {
        var listId = await GetListId();
        var requestBody = new TodoTask
        {
            Title = title,
            Categories = new List<string>
            {
                "Important",
            },
            Importance = Importance.High,
            IsReminderOn = true,
            ReminderDateTime = DateTimeTimeZone.FromDateTime(reminderTime ?? DateTime.Today, "Pacific Standard Time")
        };

        await _graphServiceClient.Me.Todo.Lists[listId].Tasks.Request().AddAsync(requestBody);
    }

    private async Task<string> GetListId()
    {
        var list = await _graphServiceClient.Me.Todo.Lists.Request().GetAsync();
        var listId = list.FirstOrDefault().Id;
        return listId;
    }
}