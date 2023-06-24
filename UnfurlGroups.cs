using System;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Azure.Identity;

namespace GraphGroupUnfurl
{
    public class UnfurlGroups
    {
        private readonly ILogger _logger;

        public UnfurlGroups(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<UnfurlGroups>();
        }

        [Function("UnfurlGroups")]
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

            if (groups?.Value is null)
            {
                _logger.LogInformation("No groups returned.");
                return;
            }

            var filteredGroups = groups.Value.Where(grp => !grp.GroupTypes.Contains("DynamicMembership")).ToList();
            foreach (var group in filteredGroups)
            {
                _logger.LogInformation($"Group: {group.DisplayName}");
                List<String> nonGroupMembers = new List<String>();

                group.Members?.ForEach(member => {
                    if (!String.IsNullOrEmpty(member.Id) && member.OdataType != "#microsoft.graph.group")
                    {
                        nonGroupMembers.Add(member.Id);
                    }
                });
                
                var nestedGroups = group?.Members?.Where(member => member.OdataType == "#microsoft.graph.group").ToList();
                bool hasNestedGroups = nestedGroups?.Count() > 0;
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
                        if (!String.IsNullOrEmpty(nestedGroupMember.Id) && nestedGroupMember.OdataType != "#microsoft.graph.group")
                        {
                            nonGroupMembers.Add(nestedGroupMember.Id);
                        }
                    });

                    _logger.LogInformation($"Group: {group.DisplayName} Non-Group Members count: {nonGroupMembers.Count()}");
                }

                var unfurledGroup = groups.Value.Where(x => x.DisplayName.Contains("UNF:"+group.DisplayName)).First();

                if (unfurledGroup is not null && hasNestedGroups)
                {
                    // Update the unfurled group membership and remove any members that are no longer in the original group
                    var unfurledMembership = unfurledGroup.Members;
                    
                    foreach (var nonGroupMember in nonGroupMembers)
                    {
                        if (!unfurledMembership.Any(x => x.Id == nonGroupMember))
                        {
                            var requestBody = new Microsoft.Graph.Models.ReferenceCreate
                            {
                                OdataId = $"https://graph.microsoft.com/v1.0/directoryObjects/{nonGroupMember}"
                            };
                            await graphServiceClient.Groups[unfurledGroup.Id].Members.Ref.PostAsync(requestBody);
                        }

                        if (unfurledMembership.Any(x => x.Id == nonGroupMember))
                        {
                            await graphServiceClient.Groups[unfurledGroup.Id].Members[nonGroupMember].Ref.DeleteAsync();
                        }
                    }

                    _logger.LogInformation($"Updated unfurled group for {group.DisplayName}");
                }
                else if (unfurledGroup is null && hasNestedGroups)
                {
                    var requestBody = new Microsoft.Graph.Models.Group
                    {
                        DisplayName = "UNF:"+group.DisplayName,
                        Description = "Unfurled group for "+group.DisplayName,
                        MailEnabled = false,
                        MailNickname = "UNF:"+group.DisplayName,
                        SecurityEnabled = true
                    };
                    var result = await graphServiceClient.Groups.PostAsync(requestBody);

                    if (result is null)
                    {
                        _logger.LogInformation($"Failed to create unfurled group for {group.DisplayName}");
                        return;
                    }

                    foreach (var nonGroupMember in nonGroupMembers)
                    {
                        var requestBody2 = new Microsoft.Graph.Models.ReferenceCreate
                        {
                            OdataId = $"https://graph.microsoft.com/v1.0/directoryObjects/{nonGroupMember}"
                        };
                        await graphServiceClient.Groups[result.Id].Members.Ref.PostAsync(requestBody2);
                    }
                    _logger.LogInformation($"Created unfurled group for {group.DisplayName}");
                }

            }


            _logger.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");
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
