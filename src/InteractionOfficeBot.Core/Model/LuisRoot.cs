﻿// <auto-generated>
// Code generated by luis:generate:cs
// Tool github: https://github.com/microsoft/botframework-cli
// Changes may cause incorrect behavior and will be lost if the code is
// regenerated.
// </auto-generated>
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.AI.Luis;
using ErrorEventArgs = Newtonsoft.Json.Serialization.ErrorEventArgs;

namespace InteractionOfficeBot.Core.Model
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
            CREATE_TODO_TASK,
            GET_ALL_TODO_TASKS,
            GET_ALL_TODO_UPCOMING_TASK,
            INSTALLED_APP_FOR_USER,
            MEMBER_CHANNEL,
            None,
            ONEDRIVE_DELETE,
            ONEDRIVE_DOWNLOAD,
            ONEDRIVE_FOLDER_CONTENTS,
            ONEDRIVE_ROOT_CONTENTS,
            ONEDRIVE_SEARCH,
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
            public DateTimeSpec[] datetime;
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

            public class _InstanceFiles
            {
                public InstanceData[] Identifier;
                public InstanceData[] Value;
            }
            public class FilesClass
            {
                public string[] Identifier;
                public string[] Value;
                [JsonProperty("$instance")]
                public _InstanceFiles _instance;
            }
            public FilesClass[] Files;

            public class _InstanceFolder
            {
                public InstanceData[] Identifier;
                public InstanceData[] Value;
            }
            public class FolderClass
            {
                public string[] Identifier;
                public string[] Value;
                [JsonProperty("$instance")]
                public _InstanceFolder _instance;
            }
            public FolderClass[] Folder;

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

            public class _InstanceReminderTime
            {
                public InstanceData[] Identifier;
                public InstanceData[] Value;
            }
            public class ReminderTimeClass
            {
                public string[] Identifier;
                public string[] Value;
                [JsonProperty("$instance")]
                public _InstanceReminderTime _instance;
            }
            public ReminderTimeClass[] ReminderTime;

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

            public class _InstanceTitle
            {
                public InstanceData[] Identifier;
                public InstanceData[] Value;
            }
            public class TitleClass
            {
                public string[] Identifier;
                public string[] Value;
                [JsonProperty("$instance")]
                public _InstanceTitle _instance;
            }
            public TitleClass[] Title;

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
                public InstanceData[] Files;
                public InstanceData[] Folder;
                public InstanceData[] Identifier;
                public InstanceData[] Message;
                public InstanceData[] ReminderTime;
                public InstanceData[] Team;
                public InstanceData[] Title;
                public InstanceData[] User;
                public InstanceData[] Value;
                public InstanceData[] datetime;
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
