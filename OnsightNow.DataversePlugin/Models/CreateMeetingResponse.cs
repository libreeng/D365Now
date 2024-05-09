using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OnsightNow.DataversePlugin.Models
{
    public class CreateMeetingResponse
    {
        [JsonProperty("id")]
        public Guid Id { get; set; }

        [JsonProperty("ownerId")]
        public Guid? OwnerId { get; set; }

        [JsonProperty("topic")]
        public string Topic { get; set; }

        [JsonProperty("allowGuests")]
        public bool AllowGuests { get; set; }

        [JsonProperty("startTime")]
        public DateTimeOffset StartTime { get; set; }

        [JsonProperty("endTime")]
        public DateTimeOffset EndTime { get; set; }

        [JsonProperty("message")]
        public string Message { get; set; }

        [JsonProperty("private")]
        public bool IsPrivate { get; set; }

        [JsonProperty("culture")]
        public string Culture { get; set; }

        [JsonProperty("participants")]
        public IEnumerable<Participant> Participants { get; set; } = Enumerable.Empty<Participant>();

        [JsonProperty("joinUrl")]
        public string JoinUrl { get; set; }

        [JsonProperty("tenantId")]
        public Guid TenantId { get; set; }


        public class Participant
        {
            [JsonProperty("type")]
            public int Type { get; set; }

            [JsonProperty("email")]
            public string Email { get; set; }

            [JsonProperty("id")]
            public Guid Id { get; set; }

            [JsonProperty("meetingId")]
            public Guid MeetingId { get; set; }

            [JsonProperty("inviteStatus")]
            public int InviteStatus { get; set; }

            [JsonProperty("tenantId")]
            public Guid TenantId { get; set; }
        }
    }
}
