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

    public async Task<List<TodoTaskEntity>> GetUpcomingTodoTasks(DateTime reminderTime)
    {
        // TODO: Get users timezone
        var userTimeZone = "Pacific Standard Time";

        var reminderTimeUtc = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(reminderTime, userTimeZone, "UTC");

		var dueDateFrom = reminderTimeUtc.ToString("u");
        var duDateTo = reminderTimeUtc.AddDays(1).ToString("u");

        var listId = await GetListId();
        var result = await _graphServiceClient
	        .Me
	        .Todo
	        .Lists[listId]
	        .Tasks
	        .Request()
	        .Filter($"dueDateTime/dateTime gt '{dueDateFrom}' and dueDateTime/dateTime lt '{duDateTo}'")
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

    public async Task CreateTodoTask(string title, DateTime reminderTime)
    {
        // TODO: Get users timezone
        var userTimeZone = "Pacific Standard Time";

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
            DueDateTime = DateTimeTimeZone.FromDateTime(reminderTime, userTimeZone)
        };

        await _graphServiceClient.Me.Todo.Lists[listId].Tasks.Request().AddAsync(requestBody);
    }

    private async Task<string?> GetListId()
    {
        var list = await _graphServiceClient.Me.Todo.Lists.Request().GetAsync();
        var listId = list.FirstOrDefault()?.Id;
        return listId;
    }
}