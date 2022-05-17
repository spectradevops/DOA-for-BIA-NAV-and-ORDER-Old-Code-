using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DOA
{
    public class EmailNotificationToApprover : CodeActivity
    {
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
                    EntityReference oppId = Opportunity.Get(executionContext);
                    EntityReference approverId = Approver.Get(executionContext);
                    EntityReference nextApprovalId = NextApproval.Get(executionContext);

                    DOAHelper helper = new DOAHelper();

                    Entity entApprover = service.Retrieve("systemuser", approverId.Id, new ColumnSet("fullname"));

                    string approver = (entApprover.Contains("fullname") ? entApprover["fullname"].ToString() : "Approver");
                    string emailbody = helper.getEmailBody(service, oppId.Id.ToString(), approver);

                    Entity entEmail = new Entity("email");
                    entEmail["subject"] = "Pending for your approval #" + context.PrimaryEntityId.ToString().ToUpper() + "#";
                    entEmail["description"] = emailbody;

                    Entity entTo = new Entity();
                    Entity entFrom = new Entity();

                    entTo = new Entity("activityparty");
                    entTo["partyid"] = new EntityReference("systemuser", entApprover.Id);
                    Entity[] entToList = { entTo };
                    entEmail["to"] = entToList;

                    Entity Queue = helper.GetResultByAttribute(service, "queue", "name", "DOA Approval", "queueid");

                    if (Queue != null)
                    {
                        entFrom = new Entity("activityparty");
                        entFrom["partyid"] = new EntityReference("queue", Queue.Id);
                        Entity[] entFromList = { entFrom };
                        entEmail["from"] = entFromList;
                    }

                    Entity oppty = helper.GetResultByAttribute(service, "opportunity", "opportunityid", oppId.Id.ToString(), "ownerid");

                    if (oppty != null)
                    {
                        Entity entcc = new Entity("activityparty");
                        entcc["partyid"] = oppty.GetAttributeValue<EntityReference>("ownerid");

                        Entity user = helper.GetResultByAttribute(service, "systemuser", "systemuserid", oppty.GetAttributeValue<EntityReference>("ownerid").Id.ToString(), "parentsystemuserid");
                        if (user.Attributes.Contains("parentsystemuserid"))
                        {
                            Entity entcc2 = new Entity("activityparty");
                            entcc2["partyid"] = user.GetAttributeValue<EntityReference>("parentsystemuserid");

                            Entity[] entccList = { entcc, entcc2 };
                            entEmail["cc"] = entccList;
                        }
                        else
                        {
                            Entity[] entccList = { entcc };
                            entEmail["cc"] = entccList;
                        }
                    }

                    entEmail["regardingobjectid"] = nextApprovalId;//new EntityReference("spectra_approval", nextApprovalId.Id);
                    Guid emailId = service.Create(entEmail);

                    //Send email
                    SendEmailRequest sendEmailReq = new SendEmailRequest()
                    {
                        EmailId = emailId,
                        IssueSend = true
                    };
                    SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);

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