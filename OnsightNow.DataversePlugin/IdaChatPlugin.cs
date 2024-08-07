using Newtonsoft.Json;
using OnsightNow.DataversePlugin.Models;
using System;
using System.Collections.Generic;

namespace OnsightNow.DataversePlugin
{
    public class IdaChatPlugin : PluginBase
    {
        public IdaChatPlugin()
            : this(string.Empty, string.Empty)
        {
        }

        public IdaChatPlugin(string unsecureConfiguration, string secureConfiguration)
            : base(typeof(IdaChatPlugin))
        {
        }

        protected override void ExecuteDataversePlugin(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }

            var context = localPluginContext.PluginExecutionContext;

            // Read Onsight NOW ClientId and ClientSecret from environment variables
            var envVars = GetEnvironmentVariables(localPluginContext);
            if (!envVars.TryGetValue("new_OnsightNowClientId", out var clientId) ||
                !envVars.TryGetValue("new_OnsightNowClientSecret", out var clientSecret))
            {
                throw new ArgumentException($"Environment variables 'new_OnsightNowClientId' and 'new_OnsightNowClientSecret' are required to use {nameof(IdaChatPlugin)}");
            }

            var nowClient = new OnsightNowClient(clientId, clientSecret);

            // Optional: allow overriding of default NOW Token API endpoint
            if (envVars.TryGetValue("new_OnsightNowTokenEndpoint", out var tokenEndpoint))
            {
                nowClient.TokenEndpoint = tokenEndpoint;
            }
            // Optional: allow overriding of default NOW Meetings API endpoint
            if (envVars.TryGetValue("new_OnsightNowIdaEndpoint", out var idaEndpoint))
            {
                nowClient.IdaChatEndpoint = idaEndpoint;
            }

            var chatInput = context.InputParameters["ChatInput"]?.ToString();
            localPluginContext.Trace($"Chat input: {chatInput}");
            var chatRequest = JsonConvert.DeserializeObject<IdaChatRequest>(chatInput);

            var idaResponse = nowClient.ChatWithIda(chatRequest);
            context.OutputParameters["ChatOutput"] = JsonConvert.SerializeObject(idaResponse);
        }
    }
}
