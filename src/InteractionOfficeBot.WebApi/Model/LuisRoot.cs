using Microsoft.Bot.Builder.AI.Luis;
using Microsoft.Bot.Builder;
using Newtonsoft.Json;
using System.Collections.Generic;
using System;
using Newtonsoft.Json.Serialization;

namespace InteractionOfficeBot.WebApi.Model
{
	 public partial class LuisRoot: IRecognizerConvert
    {
        [JsonProperty("text")]
        public string Text;

        [JsonProperty("alteredText")]
        public string AlteredText;

        public enum Intent {
            ALL_TEAMS_REQUEST,
            ALL_USER_REQUEST,
            CREATE_CHANNEL,
            CREATE_TEAM,
            INSTALLED_APP_FOR_USER,
            MEMBER_CHANNEL,
            None,
            REMOVE_CHANNEL,
            REMOVE_TEAM,
            SEND_EMAIL_TO_USER,
            SEND_MESSAGE_TO_CHANEL,
            SEND_MESSAGE_TO_USER,
            WHAT_CHANNELS_OF_TEAMS_REQUEST,
            WHO_OF_TEAMS_REQUEST
        };
        [JsonProperty("intents")]
        public Dictionary<Intent, IntentScore> Intents;

        public class _Entities
        {
            // Built-in entities
            public double[] ordinal;


            // Composites
            public class _InstanceChannel
            {
                public InstanceData[] Identifier;
                public InstanceData[] Value;
            }
            public class ChannelClass
            {
                public string[] Identifier;
                public string[] Value;
                [JsonProperty("$instance")]
                public _InstanceChannel _instance;
            }
            public ChannelClass[] Channel;

            public class _InstanceEmailSubject
            {
                public InstanceData[] Identifier;
                public InstanceData[] Value;
            }
            public class EmailSubjectClass
            {
                public string[] Identifier;
                public string[] Value;
                [JsonProperty("$instance")]
                public _InstanceEmailSubject _instance;
            }
            public EmailSubjectClass[] EmailSubject;

            public class _InstanceMessage
            {
                public InstanceData[] Identifier;
                public InstanceData[] Value;
            }
            public class MessageClass
            {
                public string[] Identifier;
                public string[] Value;
                [JsonProperty("$instance")]
                public _InstanceMessage _instance;
            }
            public MessageClass[] Message;

            public class _InstanceTeam
            {
                public InstanceData[] Identifier;
                public InstanceData[] Value;
            }
            public class TeamClass
            {
                public string[] Identifier;
                public string[] Value;
                [JsonProperty("$instance")]
                public _InstanceTeam _instance;
            }
            public TeamClass[] Team;

            public class _InstanceUser
            {
                public InstanceData[] Identifier;
                public InstanceData[] Value;
            }
            public class UserClass
            {
                public string[] Identifier;
                public string[] Value;
                [JsonProperty("$instance")]
                public _InstanceUser _instance;
            }
            public UserClass[] User;

            // Instance
            public class _Instance
            {
                public InstanceData[] Channel;
                public InstanceData[] EmailSubject;
                public InstanceData[] Identifier;
                public InstanceData[] Message;
                public InstanceData[] Team;
                public InstanceData[] User;
                public InstanceData[] Value;
                public InstanceData[] ordinal;
            }
            [JsonProperty("$instance")]
            public _Instance _instance;
        }
        [JsonProperty("entities")]
        public _Entities Entities;

        [JsonExtensionData(ReadData = true, WriteData = true)]
        public IDictionary<string, object> Properties {get; set; }

        public void Convert(dynamic result)
        {
            var app = JsonConvert.DeserializeObject<LuisRoot>(
                JsonConvert.SerializeObject(
                    result,
                    new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore, Error = OnError }
                )
            );
            Text = app.Text;
            AlteredText = app.AlteredText;
            Intents = app.Intents;
            Entities = app.Entities;
            Properties = app.Properties;
        }

        private static void OnError(object sender, ErrorEventArgs args)
        {
            // If needed, put your custom error logic here
            Console.WriteLine(args.ErrorContext.Error.Message);
            args.ErrorContext.Handled = true;
        }

        public (Intent intent, double score) TopIntent()
        {
            Intent maxIntent = Intent.None;
            var max = 0.0;
            foreach (var entry in Intents)
            {
                if (entry.Value.Score > max)
                {
                    maxIntent = entry.Key;
                    max = entry.Value.Score.Value;
                }
            }
            return (maxIntent, max);
        }
    }
}
