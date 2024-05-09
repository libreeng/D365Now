using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace OnsightNow.DataversePlugin.Models
{
    public class MeetingRequest
    {
        [JsonProperty("topic")]
        public string Topic { get; set; } = "My Onsight NOW Meeting";

        [JsonProperty("message")]
        public string Message { get; set; } = "Please join me in an Onsight NOW meeting.";

        [JsonProperty("allowGuests")]
        public bool AllowGuests { get; set; } = true;

        [JsonProperty("startTime")]
        public string StartTimeString { get; set; }

        [JsonIgnore]
        public DateTimeOffset StartTime
        {
            get => string.IsNullOrEmpty(StartTimeString) ? DateTimeOffset.MaxValue : DateTimeOffset.Parse(StartTimeString);
            set => StartTimeString = value.ToString("yyyy-MM-ddTHH:mm:sszzz");
        }

        [JsonProperty("endTime")]
        public string EndTimeString { get; set; }

        [JsonIgnore]
        public DateTimeOffset EndTime
        {
            get => string.IsNullOrEmpty(EndTimeString) ? DateTimeOffset.MaxValue : DateTimeOffset.Parse(EndTimeString);
            set => EndTimeString = value.ToString("yyyy-MM-ddTHH:mm:sszzz");
        }

        [JsonProperty("isPrivate")]
        public bool IsPrivate { get; set; }

        [JsonProperty("participants")]
        public MeetingParticipants Participants { get; set; }
    }

    public class MeetingParticipants
    {
        [JsonProperty("emails")]
        public IEnumerable<string> Emails { get; set; }

        [JsonProperty("phoneNumbers")]
        public IEnumerable<string> PhoneNumbers { get; set; }
    }
}