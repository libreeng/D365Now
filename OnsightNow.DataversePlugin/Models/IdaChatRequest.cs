using Newtonsoft.Json;
using System.Collections.Generic;

namespace OnsightNow.DataversePlugin.Models
{
    public class IdaChatRequest
    {
        [JsonProperty("variables")]
        public List<KeyValuePair<string, string>> Variables { get; set; } = new List<KeyValuePair<string, string>>();

        [JsonProperty("chatHistory")]
        public List<ChatMessage> ChatHistory { get; set; } = new List<ChatMessage>();
    }
}
