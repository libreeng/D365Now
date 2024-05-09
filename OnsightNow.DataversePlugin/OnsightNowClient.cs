using OnsightNow.DataversePlugin.Models;
using System.Collections.Generic;
using System.Net.Http;

namespace OnsightNow.DataversePlugin
{
    public class OnsightNowClient
    {
        /// <summary>
        /// OAuth scopes required to use the Onsight NOW Meetings API.
        /// </summary>
        private const string ApiScopes = "meeting_api collection_api";

        private readonly string _clientId;
        private readonly string _clientSecret;


        public OnsightNowClient(string clientId, string clientSecret)
        {
            _clientId = clientId;
            _clientSecret = clientSecret;
        }

        /// <summary>
        /// The Onsight NOW API Token endpoint. This can be overridden at runtime by the now_OnsightNowTokenEndpoint environment variable.
        /// </summary>
        public string TokenEndpoint { get; set; } = "https://login.onsightnow.com/connect/token";

        /// <summary>
        /// The Onsight NOW Meetings API endpoint. This can be overridden at runtime by the now_OnsightNowMeetingsEndpoint environment variable.
        /// </summary>
        public string MeetingsEndpoint { get; set; } = "https://api.onsightnow.com/meetings";

        /// <summary>
        /// Creates an Onsight NOW Meeting.
        /// </summary>
        /// <param name="meetingRequest"></param>
        /// <returns></returns>
        public string ScheduleMeeting(MeetingRequest meetingRequest)
        {
            var accessToken = GetAccessToken();

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", $"Bearer {accessToken}");
                using (var response = client.PostAsJson(MeetingsEndpoint, meetingRequest))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        return string.Empty;
                    }

                    var meetingResponse = response.Content.ReadAs<CreateMeetingResponse>();
                    return meetingResponse.JoinUrl;
                }
            }
        }

        /// <summary>
        /// Helper which gets an Onsight Now API Token using the client_credentials flow.
        /// </summary>
        /// <returns></returns>
        private string GetAccessToken()
        {
            var values = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("grant_type", "client_credentials"),
                new KeyValuePair<string, string>("scope", ApiScopes),
                new KeyValuePair<string, string>("client_id", _clientId),
                new KeyValuePair<string, string>("client_secret", _clientSecret)
            });

            using (var client = new HttpClient())
            {
                using (var response = client.Post(TokenEndpoint, values))
                {
                    response.EnsureSuccessStatusCode();

                    var tokenResponse = response.Content.ReadAs<OnsightNowTokenResponse>();
                    return tokenResponse.AccessToken;
                }
            }
        }
    }
}
