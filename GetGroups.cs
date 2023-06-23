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
                    List<String> nonGroupMembers = new List<String>();

                    var nestedGroups = group?.Members?.Where(member => member.OdataType == "#microsoft.graph.group").ToList();
                    
                    while (nestedGroups?.Count() > 0)
                    {
                        var nestedGroup = nestedGroups.First();
                        nestedGroups.Remove(nestedGroup);

                        var nestedGroupMembers = graphServiceClient.Groups[nestedGroup.Id]?.GetAsync((requestConfiguration) =>
                        {
                            requestConfiguration.QueryParameters.Expand = new string[] { "members($select=id,displayName)"};
                        }).Result?.Members;

                        nestedGroupMembers?.ForEach(nestedGroupMember => {
                            if (nestedGroupMember.OdataType == "#microsoft.graph.group")
                            {
                                nestedGroups.Add(nestedGroupMember);
                            }
                            else if (!String.IsNullOrEmpty(nestedGroupMember.Id))
                            {
                                    nonGroupMembers.Add(nestedGroupMember.Id);
                            }
                        });
                    }
                    _logger.LogInformation($"Group: {group.DisplayName} Non-Group Members count: {nonGroupMembers.Count()}");
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
