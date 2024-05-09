using OnsightNow.DataversePlugin.Models;
using System;

namespace OnsightNow.DataversePlugin
{
    /// <summary>
    /// The Onsight NOW Dataverse Plugin.
    /// 
    /// Plugin development guide: https://docs.microsoft.com/powerapps/developer/common-data-service/plug-ins
    /// Best practices and guidance: https://docs.microsoft.com/powerapps/developer/common-data-service/best-practices/business-logic/
    /// </summary>
    public class OnsightNowPlugin : PluginBase
    {
        public OnsightNowPlugin(string unsecureConfiguration, string secureConfiguration)
            : base(typeof(OnsightNowPlugin))
        {
            // TODO: Implement your custom configuration handling
            // https://docs.microsoft.com/powerapps/developer/common-data-service/register-plug-in#set-configuration-data
        }

        /// <summary>
        /// Main entry point for the plugin. Gets executed each time the plugin is invoked.
        /// </summary>
        /// <param name="localPluginContext"></param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
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
                throw new ArgumentException($"Environment variables 'new_OnsightNowClientId' and 'new_OnsightNowClientSecret' are required to use {nameof(OnsightNowPlugin)}");
            }

            var nowClient = new OnsightNowClient(clientId, clientSecret);

            // Optional: allow overriding of default NOW Token API endpoint
            if (envVars.TryGetValue("new_OnsightNowTokenEndpoint", out var tokenEndpoint))
            {
                nowClient.TokenEndpoint = tokenEndpoint;
            }
            // Optional: allow overriding of default NOW Meetings API endpoint
            if (envVars.TryGetValue("new_OnsightNowMeetingsEndpoint", out var meetingsEndpoint))
            {
                nowClient.MeetingsEndpoint = meetingsEndpoint;
            }

            var participantEmails = context.InputParameters["ParticipantEmails"]?.ToString()?.Split(',') ?? Array.Empty<string>();
            if (participantEmails.Length == 0)
            {
                throw new ArgumentException("Must specify at least one participant email");
            }

            // Generate Onsight NOW Meeting request using the provided participant email addresses
            localPluginContext.Trace($"Participant emails: {string.Join(", ", participantEmails)}");
            var meetingRequest = new MeetingRequest
            {
                StartTimeString = DateTimeOffset.UtcNow.ToString("g"),
                EndTimeString = DateTimeOffset.UtcNow.AddMinutes(30).ToString("g"),
                Participants = new MeetingParticipants
                {
                    Emails = participantEmails
                }
            };

            // Create the meeting in Onsight NOW; put the generated meeting URL into an output parameter
            // so that the D365 caller can use it.
            var meetingUrl = nowClient.ScheduleMeeting(meetingRequest);
            context.OutputParameters["MeetingUrl"] = meetingUrl;
        }
    }
}
