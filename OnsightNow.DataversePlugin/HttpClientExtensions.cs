using Newtonsoft.Json;
using System.Net.Http;
using System.Text;

namespace OnsightNow.DataversePlugin
{
    /// <summary>
    /// Helper methods to make async HttpClient calls synchronous. These are needed due to the way
    /// we execute the plugin within D365's Synchronous Execution Context.
    /// </summary>
    public static class HttpClientExtensions
    {
        /// <summary>
        /// Posts the given content to the specified URL. This method blocks until a response is received.
        /// </summary>
        /// <param name="client"></param>
        /// <param name="requestUri"></param>
        /// <param name="content"></param>
        /// <returns></returns>
        public static HttpResponseMessage Post(this HttpClient client, string requestUri, HttpContent content)
        {
            var task = client.PostAsync(requestUri, content);
            task.Wait();

            return task.Result;
        }

        /// <summary>
        /// Posts the given value as a JSON payload. This method blocks until a response is received.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="client"></param>
        /// <param name="requestUri"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static HttpResponseMessage PostAsJson<T>(this HttpClient client, string requestUri, T value)
        {
            var json = JsonConvert.SerializeObject(value);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var task = client.PostAsync(requestUri, content);
            task.Wait();

            return task.Result;
        }
    }
}
