using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

namespace DurableFunctionDemo4
{
    // This class implements the SharePoint webhook that gets called when a user updates a task
    public static class TaskNotifications
    {
        [FunctionName("ReceiveApproval")]
        public static async Task<object> Run(
            [HttpTrigger("OPTIONS", "POST")]HttpRequestMessage req,
            ILogger log,
            [OrchestrationClient] DurableOrchestrationClient client
            )
        {
            // Grab the validationToken URL parameter
            string validationToken = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "validationtoken", true) == 0)
                .Value;
            // If a validation token is present, we need to respond within 5 seconds by  
            // returning the given validation token. This only happens when a new 
            // web hook is being added
            if (validationToken != null)
            {
                log.LogInformation($"Validation token {validationToken} received");
                var response = req.CreateResponse(HttpStatusCode.OK);
                response.Content = new StringContent(validationToken);
                return response;
            }
            var content = await req.Content.ReadAsStringAsync();
            var notifications = JsonConvert.DeserializeObject<ResponseModel<NotificationModel>>(content).Value;
            if (notifications.Count > 0)
            {
                foreach (var notification in notifications)// we should only have one
                {
                    ChangeToken lastChangeToken = GetChangeToken(ConfigurationManager.AppSettings["taskListId"]);
                    ProcessChanges(ConfigurationManager.AppSettings["taskListId"], lastChangeToken, client,log);
                }
            }
            // if we get here we assume the request was well received
            return new HttpResponseMessage(HttpStatusCode.OK);
        }
        public static void ProcessChanges(string listID,ChangeToken lastChangeToken, DurableOrchestrationClient client,ILogger log)
        {
            using (var cc = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(ConfigurationManager.AppSettings["siteUrl"], ConfigurationManager.AppSettings["clientId"], ConfigurationManager.AppSettings["clientSecret"]))
            {
                ChangeQuery changeQuery = new ChangeQuery(false, false);
                changeQuery.Item = true;
                changeQuery.Update = true; // could handle deletes too. Just need to know if approve or reject is assumed
                changeQuery.ChangeTokenStart = lastChangeToken;
                List changedList = cc.Web.GetListById(new Guid(listID));
                var changes = changedList.GetChanges(changeQuery);
                cc.Load(changes);
                cc.ExecuteQuery();
                foreach(Change change in changes)
                   {
                    if (change is ChangeItem)
                    {
                        ListItem task = changedList.GetItemById((change as ChangeItem).ItemId);
                        ListItemVersionCollection taskVersions = task.Versions;
                        cc.Load(taskVersions);
                        cc.Load(task);
                        cc.ExecuteQuery();
                        if (taskVersions.Count < 2)
                        {
                            return;
                        }
                        var currentStatus = (string)taskVersions[0]["Status"];
                        var priorStatus = (string)taskVersions[1]["Status"];
                        string wfid = (string)task["workflowId"];
                        Console.WriteLine($"Item # ${task.Id} current status is ${currentStatus} Prior status is ${priorStatus}");
                        switch ((string)task["Action"])
                        {
                            case "DocOwnerApproval":
                                if (currentStatus != priorStatus)
                                {
                                    if (currentStatus == "Approve")
                                    {
                                        log.LogInformation("Sending event DocOwnerApproved");
                                        client.RaiseEventAsync(wfid, "DocOwnerApproved");
                                    }
                                    else
                                    {
                                        log.LogInformation("Sending event DocOwnerRejected");
                                        client.RaiseEventAsync(wfid, "DocOwnerRejected");
                                    }
                                }
                                break;
                            case "StakeHolderApproval":
                                if (currentStatus != priorStatus)
                                {
                                    if (currentStatus == "Approve")
                                    {
                                        var eventName = "StakeHolderApproval:" + ((FieldUserValue)task["AssignedTo"]).LookupId;
                                        log.LogInformation($"Sending event '${eventName}'");
                                        client.RaiseEventAsync(wfid, eventName,true);
                                    }
                                    else
                                    {
                                        log.LogInformation($"Sending event 'StakeHolderRejection'");
                                        client.RaiseEventAsync(wfid, "StakeHolderRejection");
                                    }
                                }
                                break;
                        }
                    }
                }
            };
        }
        public static ChangeToken GetChangeToken(string listID)
        {
            using (var cc = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(ConfigurationManager.AppSettings["siteUrl"], ConfigurationManager.AppSettings["clientId"], ConfigurationManager.AppSettings["clientSecret"]))
            {
                List announcementsList = cc.Web.Lists.GetByTitle("Webhook");
                ListItem item = announcementsList.GetItemById(1);
                cc.Load(item);
                cc.ExecuteQuery();
                string oldValue= (string)item["Title"];
                string newValue = string.Format("1;3;{0};{1};-1", listID, DateTime.Now.AddSeconds(-2).ToUniversalTime().Ticks.ToString());
                item["Title"] = newValue;
                item.Update();
                cc.ExecuteQuery();
                ChangeToken newChangeToken = new ChangeToken();
                newChangeToken.StringValue=oldValue;
                return newChangeToken;
            };
        }
    }
    // supporting classes required by SP  webhook
   public  class ResponseModel<T>
    {
        [JsonProperty(PropertyName = "value")]
        public List<T> Value { get; set; }
    }

    public class NotificationModel
    {
        [JsonProperty(PropertyName = "subscriptionId")]
        public string SubscriptionId { get; set; }

        [JsonProperty(PropertyName = "clientState")]
        public string ClientState { get; set; }

        [JsonProperty(PropertyName = "expirationDateTime")]
        public DateTime ExpirationDateTime { get; set; }

        [JsonProperty(PropertyName = "resource")]
        public string Resource { get; set; }

        [JsonProperty(PropertyName = "tenantId")]
        public string TenantId { get; set; }

        [JsonProperty(PropertyName = "siteUrl")]
        public string SiteUrl { get; set; }

        [JsonProperty(PropertyName = "webId")]
        public string WebId { get; set; }
    }

    public class SubscriptionModel
    {
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string Id { get; set; }

        [JsonProperty(PropertyName = "clientState", NullValueHandling = NullValueHandling.Ignore)]
        public string ClientState { get; set; }

        [JsonProperty(PropertyName = "expirationDateTime")]
        public DateTime ExpirationDateTime { get; set; }

        [JsonProperty(PropertyName = "notificationUrl")]
        public string NotificationUrl { get; set; }

        [JsonProperty(PropertyName = "resource", NullValueHandling = NullValueHandling.Ignore)]
        public string Resource { get; set; }
    }
}

