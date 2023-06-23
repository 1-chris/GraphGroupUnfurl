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
        public async void Run([TimerTrigger("0 */10 * * * *")] MyInfo myTimer)
        {
            var credential = new ChainedTokenCredential(
                new ManagedIdentityCredential(),
                new EnvironmentCredential());

            string[] scopes = new[] { "https://graph.microsoft.com/.default" };

            var graphServiceClient = new GraphServiceClient(
                credential, scopes);

            var groups = await graphServiceClient.Groups.GetAsync();

            if (groups?.Value is not null)
            {
                groups.Value.ForEach(group => {
                    _logger.LogInformation($"Group: {group.DisplayName}");
                    if (group?.Members is not null) 
                    {
                        group.Members.ForEach(member => {
                            _logger.LogInformation($"Member: {member.Id}");

                        });

                    }
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
