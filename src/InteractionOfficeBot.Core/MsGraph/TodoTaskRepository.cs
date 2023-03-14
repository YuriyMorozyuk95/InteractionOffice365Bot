using Microsoft.Graph;

namespace InteractionOfficeBot.Core.MsGraph;

public class TodoTaskRepository
{
    private readonly GraphServiceClient _graphServiceClient;

    public TodoTaskRepository(GraphServiceClient graphServiceClient)
    {
        _graphServiceClient = graphServiceClient;
    }

    public async Task<List<TodoTask>> GetTodoTasks()
    {
        var list = await _graphServiceClient.Me.Todo.Lists.Request().GetAsync();
        var listId = list.FirstOrDefault().Id;

        var result = await _graphServiceClient.Me.Todo.Lists[listId].Tasks.Request().GetAsync();
        
        return result.ToList();
    }

    //todo reminder
    public async Task CreateTodoTask(string title, DateTime? reminderTime)
    {
        var list = await _graphServiceClient.Me.Todo.Lists.Request().GetAsync();
        var listId = list.FirstOrDefault().Id;

        var requestBody = new TodoTask
        {
            Title = title,
            Categories = new List<string>
            {
                "Important",
            },
            Importance = Importance.High,
            IsReminderOn = true,
            ReminderDateTime = DateTimeTimeZone.FromDateTime(reminderTime ?? DateTime.Today)
        };

        await _graphServiceClient.Me.Todo.Lists[listId].Tasks.Request().AddAsync(requestBody);
    }
}