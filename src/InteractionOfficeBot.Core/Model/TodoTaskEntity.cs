using Microsoft.Graph;
using System.Text.Json.Serialization;
using TaskStatus = Microsoft.Graph.TaskStatus;

namespace InteractionOfficeBot.Core.Model
{
    public class TodoTaskEntity
    {
        [JsonPropertyName("title")]
        public string Title { get; set; }

        [JsonPropertyName("reminderDateTime")]
        public DateTimeTimeZone ReminderDateTime { get; set; }

        [JsonPropertyName("status")]
        public TaskStatus? Status { get; set; }
    }
}
