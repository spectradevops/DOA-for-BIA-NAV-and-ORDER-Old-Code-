using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace DOA
{
    public class EmailNotificationToApprover : CodeActivity
    {
        #region added on 24-11-2022
        [Input("ML_URL")]
        [RequiredArgument]
        public InArgument<string> ML_URL { get; set; }

        [Input("ML_Authkey")]
        [RequiredArgument]
        public InArgument<string> ML_Authkey { get; set; }

        [Input("ML_Action")]
        [RequiredArgument]
        public InArgument<string> ML_Action { get; set; }

        [Input("ML_from")]
        [RequiredArgument]
        public InArgument<string> ML_from { get; set; }
        #endregion

        [Input("Opportunity")]
        [RequiredArgument]
        [ReferenceTarget("opportunity")]
        public InArgument<EntityReference> Opportunity { get; set; }

        [Input("Approver")]
        [RequiredArgument]
        [ReferenceTarget("systemuser")]
        public InArgument<EntityReference> Approver { get; set; }

        [Input("user")]
        [RequiredArgument]
        [ReferenceTarget("systemuser")]
        public InArgument<EntityReference> User { get; set; }

        [Input("NextApproval")]
        [RequiredArgument]
        [ReferenceTarget("spectra_approval")]
        public InArgument<EntityReference> NextApproval { get; set; }

        [Input("ClientURL")]
        [RequiredArgument]
        public InArgument<string> ClientURL { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService traceService = executionContext.GetExtension<ITracingService>();
            //Obtain WorkflwoContext from the executionContext.
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            //Obtain the organization service reference.
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(User.Get(executionContext).Id);
            if (context.PrimaryEntityName.ToLower() == "spectra_approval")
            {
                if (Opportunity.Get(executionContext) != null && Approver.Get(executionContext) != null && NextApproval.Get(executionContext) != null)
                {
                    #region Email Parameters 24 Nov 2022
                    string URL = string.Empty, Authkey = string.Empty, Action = string.Empty, from1 = string.Empty, to = string.Empty, cc = string.Empty, subject = string.Empty, content = string.Empty;
                    URL = ML_URL.Get(executionContext);
                    Authkey = ML_Authkey.Get(executionContext);
                    Action = ML_Action.Get(executionContext);
                    from1 = ML_from.Get(executionContext);
                    #endregion

                    EntityReference oppId = Opportunity.Get(executionContext);
                    EntityReference approverId = Approver.Get(executionContext);
                    EntityReference nextApprovalId = NextApproval.Get(executionContext);

                    DOAHelper helper = new DOAHelper();
                    Entity entApprover = service.Retrieve("systemuser", approverId.Id, new ColumnSet("fullname"));
                    string approver = (entApprover.Contains("fullname") ? entApprover["fullname"].ToString() : "Approver");
                    string emailbody = helper.getEmailBody(service, oppId.Id.ToString(), approver);


                    #region Old Code 24 Nov 2022
                    //Entity entEmail = new Entity("email");
                    //entEmail["subject"] = "Pending for your approval #" + nextApprovalId.Id.ToString().ToUpper() + "#";
                    //entEmail["description"] = emailbody;

                    //Entity entTo = new Entity();
                    //Entity entFrom = new Entity();

                    //entTo = new Entity("activityparty");
                    //entTo["partyid"] = new EntityReference("systemuser", entApprover.Id);
                    //Entity[] entToList = { entTo };
                    //entEmail["to"] = entToList;

                    //Entity Queue = helper.GetResultByAttribute(service, "queue", "name", "DOA Approval", "queueid");

                    //if (Queue != null)
                    //{
                    //    entFrom = new Entity("activityparty");
                    //    entFrom["partyid"] = new EntityReference("queue", Queue.Id);
                    //    Entity[] entFromList = { entFrom };
                    //    entEmail["from"] = entFromList;
                    //}

                    //Entity oppty = helper.GetResultByAttribute(service, "opportunity", "opportunityid", oppId.Id.ToString(), "ownerid");

                    //if (oppty != null)
                    //{
                    //    Entity entcc = new Entity("activityparty");
                    //    entcc["partyid"] = oppty.GetAttributeValue<EntityReference>("ownerid");

                    //    Entity user = helper.GetResultByAttribute(service, "systemuser", "systemuserid", oppty.GetAttributeValue<EntityReference>("ownerid").Id.ToString(), "parentsystemuserid");
                    //    if (user.Attributes.Contains("parentsystemuserid"))
                    //    {
                    //        Entity entcc2 = new Entity("activityparty");
                    //        entcc2["partyid"] = user.GetAttributeValue<EntityReference>("parentsystemuserid");

                    //        Entity[] entccList = { entcc, entcc2 };
                    //        entEmail["cc"] = entccList;
                    //    }
                    //    else
                    //    {
                    //        Entity[] entccList = { entcc };
                    //        entEmail["cc"] = entccList;
                    //    }
                    //}

                    //entEmail["regardingobjectid"] = nextApprovalId;//new EntityReference("spectra_approval", nextApprovalId.Id);
                    //Guid emailId = service.Create(entEmail);

                    ////Send email
                    //SendEmailRequest sendEmailReq = new SendEmailRequest()
                    //{
                    //    EmailId = emailId,
                    //    IssueSend = true
                    //};
                    //SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);
                    #endregion


                    #region New Logic on 24 Nov 2022
                    subject = "Pending for your approval #" + nextApprovalId.Id.ToString().ToUpper() + "#";
                    content = emailbody.ToString();
                    string CC1 = string.Empty, CC2 = string.Empty;

                    Entity entApprover1 = service.Retrieve("systemuser", entApprover.Id, new ColumnSet("internalemailaddress"));

                    if (entApprover1.Attributes.Contains("internalemailaddress"))
                        to = entApprover1.Attributes["internalemailaddress"].ToString();

                    Entity opportunity = helper.GetResultByAttribute(service, "opportunity", "opportunityid", oppId.Id.ToString(), "ownerid");
                    Entity opportunityOwner = service.Retrieve("systemuser", opportunity.GetAttributeValue<EntityReference>("ownerid").Id, new ColumnSet("internalemailaddress", "parentsystemuserid"));

                    if (opportunityOwner.Attributes.Contains("internalemailaddress"))
                        CC1 = opportunityOwner.Attributes["internalemailaddress"].ToString();

                    if (opportunityOwner.Attributes.Contains("parentsystemuserid"))
                    {
                        Entity reportiningManager = service.Retrieve("systemuser", opportunityOwner.GetAttributeValue<EntityReference>("parentsystemuserid").Id, new ColumnSet("internalemailaddress"));
                        if (reportiningManager.Attributes.Contains("internalemailaddress"))
                            CC2 = reportiningManager.Attributes["internalemailaddress"].ToString();
                    }
                    if (CC1 != string.Empty || CC2 != string.Empty)
                    {
                        cc = CC1 + "," + CC2;
                    }
                    else
                    {
                        cc = "";
                    }

                    #region ML API Calling
                    string integrationCheck = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                              <entity name='alletech_integrationlog'>
                                                                <attribute name='alletech_name' />
                                                                <attribute name='createdon' />
                                                                <attribute name='alletech_regardingentity' />
                                                                <attribute name='alletech_integrationwith' />
                                                                <attribute name='alletech_comment' />
                                                                <attribute name='alletech_integrationlogid' />
                                                                <order attribute='createdon' descending='true' />
                                                                <filter type='and'>
                                                                  <condition attribute='alletech_entityguid' operator='eq' value='" + nextApprovalId.Id + @"' />
                                                                </filter>
                                                              </entity>
                                                            </fetch>";
                    EntityCollection integrColl = service.RetrieveMultiple(new FetchExpression(integrationCheck));
                    if (integrColl.Entities.Count == 0)
                    {
                        var result = string.Empty;
                        string json = string.Empty;
                        var httpWebRequest = (HttpWebRequest)WebRequest.Create(URL);
                        httpWebRequest.ContentType = "application/json";
                        httpWebRequest.Method = "POST";
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                        using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                        {
                            // #region CRM parameters mapping with ML parameters
                            SendingEmail app = new SendingEmail();
                            app.Authkey = Authkey;
                            app.Action = Action;
                            app.from = from1;
                            app.to = to;
                            app.cc = cc;
                            app.subject = subject;
                            app.content = content;
                            json = new JavaScriptSerializer().Serialize(app);
                            streamWriter.Write(json);
                            streamWriter.Flush();
                            streamWriter.Close();
                        }
                        //tracingService.Trace("Json generated");
                        var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                        using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                        {
                            result = streamReader.ReadToEnd();
                            if (result.Contains("success"))
                            {
                                Entity log = new Entity("alletech_integrationlog");
                                log["alletech_name"] = "Opportunity Approval Email Pushed to ML: " + nextApprovalId.Id.ToString();
                                log["alletech_integrationwith"] = new OptionSetValue(4);
                                log["alletech_regardingentity"] = "alletech_feasibility";
                                log["alletech_entityguid"] = nextApprovalId.Id.ToString();
                                log["alletech_request"] = json.ToString();
                                log["alletech_responce"] = result.ToString();
                                service.Create(log);
                            }
                            else
                            {
                                Entity log = new Entity("alletech_integrationlog");
                                log["alletech_name"] = "Opportunity Approval Email Pushed to ML: " + nextApprovalId.Id.ToString();
                                log["alletech_integrationwith"] = new OptionSetValue(4);
                                log["alletech_regardingentity"] = "alletech_feasibility";
                                log["alletech_entityguid"] = nextApprovalId.Id.ToString();
                                log["alletech_request"] = json.ToString();
                                log["alletech_responce"] = result.ToString();
                                service.Create(log);
                            }
                        }
                    }
                    #endregion
                    #endregion

                }
            }
        }

        public EntityCollection getOppProducts(IOrganizationService service, Guid oppId)
        {
            QueryExpression query = new QueryExpression("opportunityproduct");
            query.ColumnSet = new ColumnSet("extendedamount", "productid");
            query.Criteria.AddCondition(new ConditionExpression("opportunityid", ConditionOperator.Equal, oppId));
            query.Criteria.AddCondition(new ConditionExpression("spectra_approvalrequried", ConditionOperator.Equal, true));

            return service.RetrieveMultiple(query);
        }
    }
}