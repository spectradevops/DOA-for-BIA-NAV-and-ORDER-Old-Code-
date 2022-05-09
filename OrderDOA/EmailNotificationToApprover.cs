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

namespace OrderDOA
{
    public class EmailNotificationToApprover : CodeActivity
    {
        [Input("Opportunity")]
        [RequiredArgument]
        [ReferenceTarget("salesorder")]
        public InArgument<EntityReference> Order { get; set; }

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
                try
                {
                    if (Order.Get(executionContext) != null && Approver.Get(executionContext) != null && NextApproval.Get(executionContext) != null)
                    {
                        EntityReference oppId = Order.Get(executionContext);
                        EntityReference approverId = Approver.Get(executionContext);
                        EntityReference nextApprovalId = NextApproval.Get(executionContext);
                        Entity accountdetails = null;
                        Entity casedetails = null;
                        string customersegment = string.Empty;
                        string billcycle = string.Empty;


                        OrderDOAHelper helper = new OrderDOAHelper();
                        traceService.Trace("after doa helper");

                        traceService.Trace("approverId : "+ approverId.Id);
                        Entity entApprover = service.Retrieve("systemuser", approverId.Id, new ColumnSet("fullname"));

                        Entity orderdetails = service.Retrieve("salesorder", oppId.Id, new ColumnSet(true));
                        traceService.Trace("after gathering order data");
                        string arc = string.Empty;

                        if (orderdetails.Attributes.Contains("customerid"))
                        {
                            Guid accid = ((EntityReference)(orderdetails.Attributes["customerid"])).Id;
                            accountdetails = service.Retrieve("account", accid, new ColumnSet(true));

                            if (accountdetails.Attributes.Contains("name"))
                            {
                                string accountname = accountdetails.Attributes["name"].ToString();
                                EntityCollection Opportunity = getOpp(service, accountname, "opportunity", "spectra_customersegmentcode", "alletech_companynamebusiness");
                                if (Opportunity != null && Opportunity.Entities.Count > 0)
                                {
                                    Entity opp = Opportunity.Entities[0];

                                    if (opp.Attributes.Contains("spectra_customersegmentcode"))
                                    {
                                        customersegment = opp.FormattedValues["spectra_customersegmentcode"];
                                    }
                                }
                            }

                            traceService.Trace("customersegment : "+ customersegment);

                            if (orderdetails.Attributes.Contains("spectra_product"))
                            {
                                Guid prodid = ((EntityReference)orderdetails.Attributes["spectra_product"]).Id;

                                Entity Prod = service.Retrieve("product", prodid, new ColumnSet("alletech_billingcycle"));
                                if (Prod != null && Prod.Attributes.Contains("alletech_billingcycle"))
                                {
                                    billcycle = ((EntityReference)Prod.Attributes["alletech_billingcycle"]).Name;
                                }
                            }
                        }
                        traceService.Trace("billcycle : " + billcycle);

                        if (orderdetails.Attributes.Contains("spectra_caseid"))
                        {
                            casedetails = service.Retrieve("incident", orderdetails.GetAttributeValue<EntityReference>("spectra_caseid").Id, new ColumnSet(true));
                        }
                        traceService.Trace("Got case details : "+casedetails.Attributes.Count);

                        if (orderdetails.Attributes.Contains("spectra_arc"))
                        {
                           arc = orderdetails.Attributes["spectra_arc"].ToString();                           
                        }

                        string approver = (entApprover.Contains("fullname") ? entApprover["fullname"].ToString() : "Approver");

                        traceService.Trace("approver : " + approver);

                        #region Upgrade or DownGrade Check 09-May-2022
                        //Entity AccountPorduct = null;
                        //Entity OrderPorduct = null;
                        //int AccountBandwidth = 0;
                        //int OrderBandwidth = 0;
                        //int accFrequency = 0;
                        //int caseFrequency = 0;
                        //string UpgradeDowngradeStatus = string.Empty;
                        //Entity _account = (Entity)service.Retrieve("account", ((EntityReference)(orderdetails.Attributes["customerid"])).Id, new ColumnSet("alletech_businesssegment", "alletech_buildingname", "alletech_product"));
                        //if (_account.Attributes.Contains("alletech_product"))
                        //{
                        //    //  tracingService.Trace("Account is having Product");
                        //    AccountPorduct = service.Retrieve("product", ((EntityReference)_account.Attributes["alletech_product"]).Id, new ColumnSet("alletech_productsegment", "alletech_bandwidthmaster", "alletech_billingcycle"));

                        //    if (AccountPorduct.Attributes.Contains("alletech_billingcycle"))
                        //    {
                        //        Entity billingFrequency = service.Retrieve("alletech_billingcycle", ((EntityReference)AccountPorduct.Attributes["alletech_billingcycle"]).Id, new ColumnSet("alletech_monthinbillingcycle"));
                        //        if (billingFrequency.Attributes.Contains("alletech_monthinbillingcycle"))
                        //        {
                        //            accFrequency = Convert.ToInt32(billingFrequency.Attributes["alletech_monthinbillingcycle"]);
                        //        }
                        //    }
                        //    if (AccountPorduct.Attributes.Contains("alletech_bandwidthmaster"))
                        //    {
                        //        string optionText = AccountPorduct.FormattedValues["alletech_bandwidthmaster"].ToString();
                        //        string optionB = string.Empty;
                        //        for (int i = 0; i < optionText.Length; i++)
                        //        {
                        //            if (Char.IsDigit(optionText[i]))
                        //                optionB += optionText[i];
                        //        }
                        //        if (optionB.Length > 0)
                        //        {
                        //            AccountBandwidth = int.Parse(optionB);
                        //        }
                        //    }
                        //    else
                        //    {
                        //        throw new Exception("Bandwidth is empty for the current Product, please map the bandwidth on product" + ((EntityReference)_account.Attributes["alletech_product"]).Id);
                        //    }
                        //}
                        //else
                        //{
                        //    throw new Exception("Account is not associated with any of the product");
                        //}

                        //OrderPorduct = service.Retrieve("product", ((EntityReference)orderdetails.Attributes["spectra_product"]).Id, new ColumnSet("alletech_bandwidthmaster", "alletech_billingcycle"));
                        //if (OrderPorduct.Attributes.Contains("alletech_billingcycle"))
                        //{
                        //    Entity billingFrequency = service.Retrieve("alletech_billingcycle", ((EntityReference)OrderPorduct.Attributes["alletech_billingcycle"]).Id, new ColumnSet("alletech_monthinbillingcycle"));
                        //    if (billingFrequency.Attributes.Contains("alletech_monthinbillingcycle"))
                        //    {
                        //        caseFrequency = Convert.ToInt32(billingFrequency.Attributes["alletech_monthinbillingcycle"]);
                        //    }
                        //}
                        //if (OrderPorduct.Attributes.Contains("alletech_bandwidthmaster"))
                        //{
                        //    string optionText = OrderPorduct.FormattedValues["alletech_bandwidthmaster"].ToString();
                        //    string optionB = string.Empty;
                        //    for (int i = 0; i < optionText.Length; i++)
                        //    {
                        //        if (Char.IsDigit(optionText[i]))
                        //            optionB += optionText[i];
                        //    }
                        //    if (optionB.Length > 0)
                        //    {
                        //        OrderBandwidth = int.Parse(optionB);
                        //    }
                        //}
                        //else
                        //{
                        //    throw new Exception("Bandwidth is empty for the Upgraded Product, please map the bandwidth on product");
                        //}

                        //if (OrderBandwidth < AccountBandwidth)
                        //{
                        //    UpgradeDowngradeStatus = "Downgrade";
                        //}
                        //else if (OrderBandwidth > AccountBandwidth)
                        //{
                        //    UpgradeDowngradeStatus = "Upgrade";
                        //}
                        //else
                        //{
                        //    if (OrderBandwidth == AccountBandwidth)
                        //    {
                        //        if (caseFrequency < accFrequency)
                        //        {
                        //            UpgradeDowngradeStatus = "Downgrade";
                        //        }
                        //        else
                        //        {
                        //            UpgradeDowngradeStatus = "Upgrade";
                        //        }
                        //    }
                        //}

                        #endregion Ends here

                        string emailbody = helper.getEmailBody(service, approver, orderdetails, accountdetails, casedetails, customersegment, billcycle,arc, traceService);

                        Entity entEmail = new Entity("email");
                        entEmail["subject"] = "Pending for your approval";
                        entEmail["description"] = emailbody;

                        //To
                        Entity Temp = new Entity();
                        Temp = new Entity("activityparty");
                        Temp["partyid"] = new EntityReference("systemuser", entApprover.Id);
                        Entity[] entToList = { Temp };
                        entEmail["to"] = entToList;

                        #region CC
                        List<Entity> entccList = new List<Entity>();
                        //if (accountdetails.Attributes.Contains("spectra_servicerelationshipmanagerid"))
                        //{
                        //    Temp = new Entity("activityparty");
                        //    Temp["partyid"] = accountdetails.GetAttributeValue<EntityReference>("spectra_servicerelationshipmanagerid");

                        //    entccList.Add(Temp);
                        //}

                        EntityCollection CCcoll = helper.getApprovalConfig(service, "CC", "B2BUP_");
                        if(CCcoll.Entities.Count>0)
                        {
                            if (CCcoll.Entities[0].Contains("spectra_approver"))
                            {
                                Temp = new Entity("activityparty");
                                Temp["partyid"] =CCcoll.Entities[0].GetAttributeValue<EntityReference>("spectra_approver");

                                entccList.Add(Temp);
                            }
                        }

                        //Temp = new Entity("activityparty");
                        //Temp["partyid"] = accountdetails.GetAttributeValue<EntityReference>("ownerid");
                        //entccList.Add(Temp);

                        //Entity user = helper.GetResultByAttribute(service, "systemuser", "systemuserid", accountdetails.GetAttributeValue<EntityReference>("ownerid").Id.ToString(), "parentsystemuserid");
                        //if (user.Attributes.Contains("parentsystemuserid"))
                        //{
                        //    Temp = new Entity("activityparty");
                        //    Temp["partyid"] = user.GetAttributeValue<EntityReference>("parentsystemuserid");

                        //    entccList.Add(Temp);
                        //}

                        entEmail["cc"] = entccList.ToArray();
                        #endregion

                        //from
                        Entity Queue = helper.GetResultByAttribute(service, "queue", "name", "DOA Approval", "queueid");
                        if (Queue != null)
                        {
                            Temp = new Entity("activityparty");
                            Temp["partyid"] = new EntityReference("queue", Queue.Id);
                            Entity[] entFromList = { Temp };
                            entEmail["from"] = entFromList;
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
                catch (Exception ex) {
                    throw new Exception("Error From EmailNotificationToApprover"+ex.Message);
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

        public EntityCollection getOpp(IOrganizationService service, string accname,string entityname,string columns,string fieldname)
        {
            QueryExpression query = new QueryExpression(entityname);
            query.ColumnSet = new ColumnSet(columns);
            query.Criteria.AddCondition(new ConditionExpression(fieldname, ConditionOperator.Equal, accname));

            return service.RetrieveMultiple(query);
        }

    }
}

