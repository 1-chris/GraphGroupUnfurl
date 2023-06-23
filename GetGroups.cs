using System;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Azure.Identity;

namespace GraphGroupUnfurl
{
    public class GetGroups
    {
        private readonly ILogger _logger;

        public GetGroups(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<GetGroups>();
        }

        [Function("GetGroups")]
        public async Task Run([TimerTrigger("0 */1 * * * *")] MyInfo myTimer)
        {
            var credential = new ChainedTokenCredential(
                new ManagedIdentityCredential(),
                new EnvironmentCredential());

            string[] scopes = new[] { "https://graph.microsoft.com/.default" };

            var graphServiceClient = new GraphServiceClient(
                credential, scopes);

            var groups = await graphServiceClient.Groups.GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Expand = new string[] { "members($select=id,displayName)"};
            });

            if (groups?.Value is not null)
            {
                groups.Value.ForEach(group => {
                    _logger.LogInformation($"Group: {group.DisplayName}");

                    group?.Members?.ForEach(async member => {
                        _logger.LogInformation($"Group: {group.DisplayName} Member: {member.Id} Type: {member.OdataType}");
                        if (member.OdataType == "#microsoft.graph.group")
                        {
                            var subGroupMembers = await graphServiceClient.Groups[$"{member.Id}"].GetAsync((requestConfiguration) =>
                            {
                                requestConfiguration.QueryParameters.Expand = new string[] { "members($select=id,displayName)"};
                            });
                            subGroupMembers?.Members?.ForEach(subMember => {
                                _logger.LogInformation($"Group: {group.DisplayName} Member: {member.Id} Type: {member.OdataType}");
                            });
                        }
                    });

                });
            }

            _logger.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");
            _logger.LogInformation($"Next timer schedule at: {myTimer.ScheduleStatus.Next}");
        }
    }

    public class MyInfo
    {
        public MyScheduleStatus ScheduleStatus { get; set; } 

        public bool IsPastDue { get; set; }
    }

    public class MyScheduleStatus
    {
        public DateTime Last { get; set; }

        public DateTime Next { get; set; }

        public DateTime LastUpdated { get; set; }
    }
}
