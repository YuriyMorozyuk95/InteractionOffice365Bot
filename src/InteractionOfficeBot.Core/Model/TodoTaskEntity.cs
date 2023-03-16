using Microsoft.Graph;
using TaskStatus = Microsoft.Graph.TaskStatus;

namespace InteractionOfficeBot.Core.Model
{
    public class TodoTaskEntity
    {
        public string? Title { get; set; }

        public DateTime? ReminderDateTime { get; set; }

        public TaskStatus? Status { get; set; }
    }
}
