using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using System.Configuration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using Newtonsoft.Json;
using Microsoft.Extensions.Logging;
using System.Threading;
using ExecutionContext = Microsoft.Azure.WebJobs.ExecutionContext;
//using OfficeDevPnP;

namespace DurableFunctionDemo4
{
    public class Orchestrator
    {
        /// <summary>
        /// This is the data structure that will be returned from GetListItemData
        /// </summary>
        public class ListItemData
        {
            public string Title;
            public int DocumentOwner; //the lookupID
            public List<int> StakeHolders;
        }
        /// <summary>
        /// This is the data structure that will be passed to ScheduleStakeHolderApproval
        /// </summary>
        public class ScheduleStakeHolderApprovalParams
        {

            public int stakeHolderId;
            public string instanceId;
        }
        /// <summary>
        /// This is the data structure that will be passed into the Publish function
        /// </summary>
        public class ApprovalStartInfo
        {
            public string startedByEmail;
            public int itemId;
        }
        //public class DocOwnerApprovalInfo
        //{
        //    string ApprovalStatus;
        //    string instanceID;
        //}
        /// <summary>
        /// This is the main Orchestartor function. it is responsible for coordinating the work of all the Actibity Functions
        /// </summary>
        /// <param name="context">A DurableOrchestrationContext supplied by the runtime </param>
        /// <param name="log">An ILogger provided by the runtime</param>
        /// <returns></returns>

        [FunctionName("Publish")]
        public static async Task Publish(
            [OrchestrationTrigger] DurableOrchestrationContext context,
            ILogger log
            )

        {
            // get the parameters passed in to the this (the publish/orchestrator function)
            ApprovalStartInfo approvalStartInfo = context.GetInput<ApprovalStartInfo>();
         

            context.SetCustomStatus("Fetching list data");

            // 'Call' a function to fetch the info from the sharepoint item being approvd.
            // this call writes an entry to the workitem queue and stops execution of the the publish/orchestrator function)
            // the GetListItemData function gets triggered when it sees the entry in the workItem queue
            ListItemData listItemData = await context.CallActivityAsync<ListItemData>("GetListItemData", approvalStartInfo);

            // when the GetListItemData returns a value , that value gets written into the controle queuw
            // that will cause this function to restart. When this function restarts, and comes to the line above,
            // the CallActivityAsync method will see that GetListItemData has already been called  (by looking in the control queue
            // So it wont call the function again, it will just put the returned valuse in the listItemData variable.

            context.SetCustomStatus("Sceduling Docownwer Approval");

            // 'Call' a function to get the document ownwers approval
            // this call writes an entry to the workitem queue and stops execution of the the publish/orchestrator function)
            // the ScheduleDocOwnerApproval function gets triggered when it sees the entry in the workItem queue
            // the ScheduleDocOwnerApproval function creates a task for the document owner to approve/reject.
            // when the document ownwer approves or rejects the task, the TaskNotifications Azure Function will send a 
            // DocOwnerApproved or  DocOwnerRejected external event to this function.
            await context.CallActivityAsync<string>("ScheduleDocOwnerApproval", new ScheduleDocOwnerApprovalParms() { instanceId = context.InstanceId, DocumentOwner = listItemData.DocumentOwner});
            var docOwnerApproved = context.WaitForExternalEvent("DocOwnerApproved");
            var docOwnerRejected = context.WaitForExternalEvent("DocOwnerRejected");
            context.SetCustomStatus("Awaiting  Docownwer approval");

            // wait for one of those two events (no tiemout here)
            var winner = await Task.WhenAny(docOwnerApproved, docOwnerRejected);
            if (winner == docOwnerRejected)
            {
                // if the Document Owner rejected the task trigger execution of the DocownerRejected function
                // which just sends an email. The workflow is then done!
                context.SetCustomStatus("Docownwer Rejected");
                await context.CallActivityAsync<string>("DocownerRejected", approvalStartInfo);
            }
            else
            {
                // stakeholder IDs is an array of the ids of the stakeholders (fron the user information list)
                List<int> stakeholderIDs = listItemData.StakeHolders;
                
                // create a list of tasks that need to be approved -- one per stakeholder
                var stakeHolderApprovalTasks = new Task[stakeholderIDs.Count];

                for (int i = 0; i < stakeholderIDs.Count; i++)
                {
                    //'Call' the ScheduleStakeholderApproval function to get the stakeholders approval.
                    // The ScheduleStakeholderApproval function will create a task in the Tasks list for the stakeholder.
                    // When the stakeholder Approves the task, the TaskNotifications Azure Function will send a 
                    // StakeHolderApproval:xx (where xx is the ID of the stakeholder) event to this function.
                    // When the stakeholder Rejects  the task, the TaskNotifications Azure Function will send a 
                    // StakeHolderRejection event to this function. 
                    await context.CallActivityAsync("ScheduleStakeholderApproval", new ScheduleStakeHolderApprovalParams() { instanceId = context.InstanceId, stakeHolderId = stakeholderIDs[i] });
                    string eventName = "StakeHolderApproval:" + stakeholderIDs[i];

                    // add the approval task to the list of stakeholder tasks that we will be waiting on/
                    stakeHolderApprovalTasks[i] = context.WaitForExternalEvent(eventName);
                }

                // the stakeHolderApprovalTask will be completed when ALL stakeholders approve their tasks
                var stakeHolderApprovalTask = Task.WhenAll(stakeHolderApprovalTasks);
                // the stakeHolderRejectionTask will be completed when ANY stakeholder rejects her task
                var stakeHolderRejectionTask = context.WaitForExternalEvent("StakeHolderRejection");


                // Now we need to set up a CancellationTokenSource so that we can wait for a certain amout of ti,e
                using (var cts = new CancellationTokenSource())
                {
                    //for this sample we will wait for just one minute.
                    var timeoutAt = context.CurrentUtcDateTime.AddMinutes(1);
                    //now we can create a task that will timeout after one minute.
                    var timeoutTask = context.CreateTimer(timeoutAt, cts.Token);

                    // And now.... wait for EITHER the timeout, a SINGLE rejection, or ALL stakeholders to approve.
                    var stakeHolderResults = await Task.WhenAny(stakeHolderApprovalTask, timeoutTask, stakeHolderRejectionTask);

                    if (stakeHolderResults == stakeHolderApprovalTask)
                    {
                        cts.Cancel(); //  cancel the timeout task
                        log.LogCritical("received stakeholder approvals");
                        // copy the file 
                        await context.CallActivityAsync<ListItemData>("CopyFile", approvalStartInfo);

                    }
                    else if (stakeHolderResults == stakeHolderRejectionTask)
                    {
                        cts.Cancel(); // we should cancel the timeout task
                        log.LogCritical("A stakeholder rejected");
                        // should send email that action was rejected
                    }
                    else
                    {
                        // timed out
                        log.LogCritical(" Timed out waiting for stakeholder approvals");
                        // copy the file 
                        await context.CallActivityAsync<ListItemData>("CopyFile", approvalStartInfo);


                    }
                }
            }
          
        }
        [FunctionName("GetListItemData")]
        public static ListItemData GetListItemData(
            [ActivityTrigger] ApprovalStartInfo approvalStartInfo,
            ILogger log,
            ExecutionContext context)
        {

            using (var cc = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(ConfigurationManager.AppSettings["siteUrl"], ConfigurationManager.AppSettings["clientId"], ConfigurationManager.AppSettings["clientSecret"]))

            {
                List docLib = cc.Web.Lists.GetByTitle("Drafts");
                ListItem item = docLib.GetItemById(approvalStartInfo.itemId);
                cc.Load(item);
                cc.ExecuteQuery();
                var listItemData = new ListItemData();
                listItemData.DocumentOwner = ((FieldUserValue)item["DocumentOwner"]).LookupId;
                listItemData.StakeHolders = new List<int>();
                foreach (FieldUserValue sh in (FieldUserValue[])item["StakeHolders"])
                {
                    listItemData.StakeHolders.Add(sh.LookupId);
                }
                return listItemData;
            };
        }
        [FunctionName("CopyFile")]
        public static void CopyFile(
       [ActivityTrigger] ApprovalStartInfo approvalStartInfo,
       ILogger log,
       ExecutionContext context)
        {
            using (var cc = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(ConfigurationManager.AppSettings["siteUrl"], ConfigurationManager.AppSettings["clientId"], ConfigurationManager.AppSettings["clientSecret"]))
            {
                List docLib = cc.Web.Lists.GetByTitle("Drafts");
                ListItem item = docLib.GetItemById(approvalStartInfo.itemId);
                Microsoft.SharePoint.Client.File file = item.File;

                cc.Load(file);
                cc.Load(item);
                
                cc.ExecuteQuery();
                string dest = ConfigurationManager.AppSettings["siteUrl"] + "Published/" + file.Name;
                file.CopyTo(dest,true);
                cc.ExecuteQuery();

            };
        }

        public class ScheduleDocOwnerApprovalParms
        {
            public int DocumentOwner;
            public string instanceId;
        }
       [FunctionName("ScheduleDocOwnerApproval")]
        public static void ScheduleDocOwnerApproval(
            [ActivityTrigger] ScheduleDocOwnerApprovalParms parms,
            TraceWriter log,
            ExecutionContext context
          )
        {
            using (var cc = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(ConfigurationManager.AppSettings["siteUrl"], ConfigurationManager.AppSettings["clientId"], ConfigurationManager.AppSettings["clientSecret"]))
            {
                List taskList = cc.Web.Lists.GetByTitle("Tasks");
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem taskItem = taskList.AddItem(itemCreateInfo);
                taskItem["Title"] = "Approval required";
                taskItem["AssignedTo"] = new FieldUserValue() { LookupId = parms.DocumentOwner };
                taskItem["workflowId"] = parms.instanceId;
                taskItem["Action"] = "DocOwnerApproval";
                taskItem.Update();
                cc.ExecuteQuery();
                var results = new ListItemData();
            };
        }

        [FunctionName("DocownerRejected")]
        public static void DocownerRejected(
            [ActivityTrigger] ApprovalStartInfo approvalStartInfo,
            TraceWriter log,
            ExecutionContext context
       )
        {
            
            using (var cc = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(ConfigurationManager.AppSettings["siteUrl"], ConfigurationManager.AppSettings["clientId"], ConfigurationManager.AppSettings["clientSecret"]))
            {
                var emailp = new EmailProperties();
                emailp.BCC = new List<string> { approvalStartInfo.startedByEmail };
                emailp.To = new List<string> { approvalStartInfo.startedByEmail };
                emailp.From = "from@mail.com";
                emailp.Body = "<b>rejected</b>";
                emailp.Subject = "rejected";

                Utility.SendEmail(cc, emailp);
                cc.ExecuteQuery();
            };
        }



        [FunctionName("ScheduleStakeholderApproval")]

        public static void ScheduleStakeholderApproval(
            [ActivityTrigger] ScheduleStakeHolderApprovalParams parms, // lokuupID of the user
            TraceWriter log,
            ExecutionContext context
            )
        {

            using (var cc = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(ConfigurationManager.AppSettings["siteUrl"], ConfigurationManager.AppSettings["clientId"], ConfigurationManager.AppSettings["clientSecret"]))
            {
                List taskList = cc.Web.Lists.GetByTitle("Tasks");
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem taskItem = taskList.AddItem(itemCreateInfo);
                taskItem["Title"] = "StakeHolder Approval required";
                taskItem["AssignedTo"] = new FieldUserValue() { LookupId = parms.stakeHolderId };
                taskItem["workflowId"] = parms.instanceId;
                taskItem["Action"] = "StakeHolderApproval";
                taskItem.Update();
                cc.ExecuteQuery();
                var results = new ListItemData();



            }
        }
        [FunctionName("ApprovalStart")]
        public static async Task<HttpResponseMessage> ApprovalStart(
            [HttpTrigger(AuthorizationLevel.Function, "OPTIONS", "POST")]HttpRequestMessage req,
            [OrchestrationClient]DurableOrchestrationClient starter,
            TraceWriter log,
              ExecutionContext context)
        {
            // Function input comes from the request content.

            if (req.Method == HttpMethod.Options)
            {
                var response = req.CreateResponse();
                if (AddCORSHeaders(req, ref response, "POST,OPTIONS", log))
                {
                    response.StatusCode = HttpStatusCode.OK;
                }
                else
                {
                    response.StatusCode = HttpStatusCode.InternalServerError;
                }
                return response;
            }
            ApprovalStartInfo approvalStartInfo = await req.Content.ReadAsAsync<ApprovalStartInfo>();
            string instanceId = await starter.StartNewAsync("Publish", approvalStartInfo);
            log.Info($"Started orchestration with ID = '{instanceId}'.");
            var resp = starter.CreateCheckStatusResponse(req, instanceId);
            if (AddCORSHeaders(req, ref resp, "POST,OPTIONS", log))
            {
                resp.StatusCode = HttpStatusCode.OK;
            }
            else
            {
                resp.StatusCode = HttpStatusCode.InternalServerError;
            }
            return resp;

        }


        [FunctionName("GetAllStatus")]
        public static async Task<HttpResponseMessage> GetAllStatus(
                [HttpTrigger(AuthorizationLevel.Anonymous, "get")]HttpRequestMessage req,
                [OrchestrationClient] DurableOrchestrationClient client,
                TraceWriter log)
        {
            IList<DurableOrchestrationStatus> instances = await client.GetStatusAsync(); // You can pass CancellationToken as a parameter.
            foreach (var instance in instances)
            {
                log.Info(JsonConvert.SerializeObject(instance));
            };
            var resp = req.CreateResponse(HttpStatusCode.OK);
            resp.Content = new StringContent(JsonConvert.SerializeObject(instances));
            AddCORSHeaders(req, ref resp, "GET", log);
            return resp;
        }
        [FunctionName("TerminateInstance")]
        public static async Task<HttpResponseMessage> TerminateInstanceAsync(
            [OrchestrationClient] DurableOrchestrationClient client,
            [HttpTrigger(AuthorizationLevel.Function, "PUT", "OPTIONS", Route = "TerminateInstance/{stringId}")]HttpRequestMessage request,
             TraceWriter log,
            string stringId)
        {
            string reason = "canceled";
            await client.TerminateAsync(stringId, reason);
            var resp = request.CreateResponse(HttpStatusCode.OK);
            AddCORSHeaders(request, ref resp, "PUT", log);
            return resp;

        }

        [FunctionName("PurgeHistory")]
        public static async Task<HttpResponseMessage> PurgeHistory(
         [OrchestrationClient] DurableOrchestrationClient client,
         [HttpTrigger(AuthorizationLevel.Function, "PUT", "OPTIONS", Route = "PurgeHistory/{stringId}")]HttpRequestMessage request,
          TraceWriter log,
         string stringId)
        {
            string reason = "canceled";
            await client.PurgeInstanceHistoryAsync(stringId);
            var resp = request.CreateResponse(HttpStatusCode.OK);
            AddCORSHeaders(request, ref resp, "PUT", log);
            return resp;


        }


        public static bool AddCORSHeaders(HttpRequestMessage req, ref HttpResponseMessage resp, string verbs, TraceWriter log)
        {

            if (req.Headers.Contains("Origin"))
            {
                var origin = req.Headers.GetValues("Origin").FirstOrDefault();
                if (origin != null)
                {
                    List<string> AllowedDomains = ConfigurationManager.AppSettings["AllowedOrigins"].Split(',').ToList();
                    if (AllowedDomains.Contains(origin))
                    {
                        resp.Headers.Add("Access-Control-Allow-Credentials", "true");
                        resp.Headers.Add("Access-Control-Allow-Origin", origin);
                        resp.Headers.Add("Access-Control-Allow-Methods", verbs);
                        resp.Headers.Add("Access-Control-Allow-Headers", "Content-Type");

                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return true;
            }



        }




    }
}
