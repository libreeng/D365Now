using Newtonsoft.Json;

namespace OnsightNow.DataversePlugin.Models
{
    internal class OnsightNowTokenResponse
    {
        [JsonProperty("access_token")]
        public string AccessToken { get; set; }
    }
}
