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
    public class CreateApprovals : CodeActivity
    {
        [Input("City")]
        [RequiredArgument]
        [ReferenceTarget("alletech_city")]
        public InArgument<EntityReference> City { get; set; }

        [Input("user")]
        [RequiredArgument]
        [ReferenceTarget("systemuser")]
        public InArgument<EntityReference> User { get; set; }

        [Input("ClientURL")]
        [RequiredArgument]
        public InArgument<string> ClientURL { get; set; }

        [Output("ApprovalCreated")]
        public OutArgument<bool> ApprovalCreated { get; set; }

        string searchData = "_IPADDRESS_";

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService traceService = executionContext.GetExtension<ITracingService>();
            //Obtain WorkflwoContext from the executionContext.
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            //Obtain the organization service reference.
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            //IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            IOrganizationService service = serviceFactory.CreateOrganizationService(User.Get(executionContext).Id);

            if (context.PrimaryEntityName.ToLower() == "opportunity")
            {

                Dictionary<int, Guid> approvals = new Dictionary<int, Guid>();

                //updated for City Head Change
                Entity OpportunityDetails = service.Retrieve("opportunity", context.PrimaryEntityId, new ColumnSet(true));
                //throw new Exception(""+context.PrimaryEntityId+" "+context.PrimaryEntityName);
                if (OpportunityDetails.Attributes.Contains("ownerid"))
                {

                    traceService.Trace("Inside IF");
                    #region New Changes done on 16-Sep-2020 by VLABS
                    Entity oppowner = service.Retrieve("systemuser", ((EntityReference)OpportunityDetails.Attributes["ownerid"]).Id, new ColumnSet("spectra_cityhead"));
                    if (oppowner.Attributes.Contains("spectra_cityhead"))
                    {
                        approvals.Add(0, ((EntityReference)oppowner["spectra_cityhead"]).Id);
                    }
                    else
                        throw new InvalidPluginExecutionException("City Head is not maaped with Opportunity owner, please contact System Administrator");
                    #endregion


                    #region Old code which was used till 15-Sep-2020

                    //Entity oppowner = service.Retrieve("systemuser", ((EntityReference)OpportunityDetails.Attributes["ownerid"]).Id, new ColumnSet("alletech_operatingcity"));
                    //if (oppowner.Attributes.Contains("alletech_operatingcity"))
                    //{
                    //    Entity entCity = service.Retrieve("alletech_city", ((EntityReference)oppowner.Attributes["alletech_operatingcity"]).Id, new ColumnSet("spectra_cityheadid"));
                    //    if (entCity.Contains("spectra_cityheadid"))
                    //    {
                    //        approvals.Add(0, ((EntityReference)entCity["spectra_cityheadid"]).Id);
                    //    }
                    //    else
                    //        throw new InvalidPluginExecutionException("City Head for the selected city is null, please contact System Administrator");
                    //}
                    //else
                    //{
                    //    traceService.Trace("Inside else");
                    //    //throw new Exception("else part");
                    //    if (City.Get(executionContext).Id != null)
                    //    {
                    //        Entity entCity = service.Retrieve("alletech_city", City.Get(executionContext).Id, new ColumnSet("spectra_cityheadid"));
                    //        if (entCity.Contains("spectra_cityheadid"))
                    //        {
                    //            approvals.Add(0, ((EntityReference)entCity["spectra_cityheadid"]).Id);
                    //        }
                    //        else
                    //            throw new InvalidPluginExecutionException("City Head for the selected city is null, please contact System Administrator");
                    //    }
                    //}
                    #endregion OLD code is end
                }

                EntityReference prodseg = OpportunityDetails.GetAttributeValue<EntityReference>("alletech_productsegment");

                EntityCollection entCollOppProd = getOppProducts(service, context.PrimaryEntityId, false);

                foreach (Entity entOppProd in entCollOppProd.Entities)
                {
                    if (entOppProd.Contains("spectra_approvalrequried") && (bool)entOppProd["spectra_approvalrequried"])
                    {
                        if (entOppProd.Contains("productid"))
                        {
                            EntityReference prodId = (EntityReference)entOppProd["productid"];

                            Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_businesssegmentlookup", "alletech_plantype", "alletech_chargetype", "alletech_grossplaninvoicevalueinr"));
                            if (entProd.Contains("alletech_businesssegmentlookup") && ((EntityReference)entProd["alletech_businesssegmentlookup"]).Name.ToLower() == "business")
                            {
                                decimal percentAge = 0;
                                decimal extendedAmt = ((Money)entOppProd["extendedamount"]).Value;
                                int chargetype = entProd.GetAttributeValue<OptionSetValue>("alletech_chargetype").Value;
                                int plantype = entProd.GetAttributeValue<OptionSetValue>("alletech_plantype").Value;

                                if (plantype == 569480002 && chargetype == 569480001) //ip address
                                {
                                    var cnt = prodId.Name.Substring(prodId.Name.IndexOf(searchData) + searchData.Length);
                                    cnt = Regex.Replace(cnt, "[^0-9]+", string.Empty);
                                    Int64 count = Convert.ToInt64(cnt);
                                    decimal price = ((Money)entOppProd["priceperunit"]).Value;

                                    if (extendedAmt < 200)
                                    {
                                        percentAge = extendedAmt;
                                        EntityCollection entCollAppConfig = getApprovalConfig(service, "IPADDRESS", null);//, percentAge);
                                        foreach (Entity entAppConfig in entCollAppConfig.Entities)
                                        {
                                            if ((entAppConfig.Contains("spectra_quantity") && count >= Convert.ToInt64(entAppConfig["spectra_quantity"])) || !entAppConfig.Contains("spectra_quantity"))
                                            {
                                                if (entAppConfig.Contains("spectra_minpercentage") && entAppConfig.Contains("spectra_maxpercentage")
                                                    && (percentAge >= ((decimal)entAppConfig["spectra_minpercentage"])) && (percentAge <= ((decimal)entAppConfig["spectra_maxpercentage"])))
                                                {
                                                    if (entAppConfig.Contains("spectra_orderby") && entAppConfig.Contains("spectra_approver"))
                                                    {
                                                        KeyValuePair<int, Guid> apprDuplicate = approvals.FirstOrDefault(a => a.Key == Convert.ToInt16(entAppConfig.FormattedValues["spectra_orderby"]));
                                                        if (apprDuplicate.Equals(new KeyValuePair<int, Guid>()))
                                                        {
                                                            //approvals.Add(Convert.ToInt16(entAppConfig.FormattedValues["spectra_orderby"]), ((EntityReference)entAppConfig["spectra_approver"]).Id);
                                                            KeyValuePair<int, Guid> apprDuplicateCond2 = approvals.FirstOrDefault(a => a.Value == ((EntityReference)entAppConfig["spectra_approver"]).Id);
                                                            if (apprDuplicateCond2.Equals(new KeyValuePair<int, Guid>()))
                                                                approvals.Add(Convert.ToInt16(entAppConfig.FormattedValues["spectra_orderby"]), ((EntityReference)entAppConfig["spectra_approver"]).Id);
                                                            else
                                                            {
                                                                approvals.Remove(apprDuplicateCond2.Key);
                                                                approvals.Add(Convert.ToInt16(entAppConfig.FormattedValues["spectra_orderby"]), ((EntityReference)entAppConfig["spectra_approver"]).Id);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                //RC || OTC
                                else if (chargetype == 569480001 || chargetype == 569480002)// 
                                {
                                    if (entProd.Contains("alletech_grossplaninvoicevalueinr"))
                                    {
                                        decimal floorDisc = ((Money)entProd["alletech_grossplaninvoicevalueinr"]).Value;
                                        if (extendedAmt < floorDisc)
                                        {
                                            percentAge = (floorDisc - extendedAmt) / floorDisc * 100;
                                            percentAge = decimal.Round(percentAge, 2, MidpointRounding.AwayFromZero);

                                            EntityCollection entCollAppConfig = getApprovalConfig(service, (chargetype == 569480001 ? "RC" : "OTC"), prodseg.Id.ToString());//, percentAge);
                                            foreach (Entity entAppConfig in entCollAppConfig.Entities)
                                            {
                                                if (entAppConfig.Contains("spectra_minpercentage") && entAppConfig.Contains("spectra_maxpercentage")
                                                    && (percentAge >= ((decimal)entAppConfig["spectra_minpercentage"])) && (percentAge <= ((decimal)entAppConfig["spectra_maxpercentage"])))
                                                {
                                                    if (entAppConfig.Contains("spectra_orderby") && entAppConfig.Contains("spectra_approver"))
                                                    {
                                                        KeyValuePair<int, Guid> apprDuplicate = approvals.FirstOrDefault(a => a.Key == Convert.ToInt16(entAppConfig.FormattedValues["spectra_orderby"]));
                                                        if (apprDuplicate.Equals(new KeyValuePair<int, Guid>()))
                                                        {
                                                            //approvals.Add(Convert.ToInt16(entAppConfig.FormattedValues["spectra_orderby"]), ((EntityReference)entAppConfig["spectra_approver"]).Id);
                                                            KeyValuePair<int, Guid> apprDuplicateCond2 = approvals.FirstOrDefault(a => a.Value == ((EntityReference)entAppConfig["spectra_approver"]).Id);
                                                            if (apprDuplicateCond2.Equals(new KeyValuePair<int, Guid>()))
                                                                approvals.Add(Convert.ToInt16(entAppConfig.FormattedValues["spectra_orderby"]), ((EntityReference)entAppConfig["spectra_approver"]).Id);
                                                            else
                                                            {
                                                                approvals.Remove(apprDuplicateCond2.Key);
                                                                approvals.Add(Convert.ToInt16(entAppConfig.FormattedValues["spectra_orderby"]), ((EntityReference)entAppConfig["spectra_approver"]).Id);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                //approvals = approvals.OrderByDescending(a => a.Key).;
                Guid approvalId = Guid.Empty;
                string optionValuePreFix = "11126000";
                int firstApproval = 0;
                if (approvals.Count > 0)
                {
                    var approvalSorted = from appr in approvals
                                         orderby appr.Key ascending
                                         select appr;
                    firstApproval = approvalSorted.First().Key;
                }
                foreach (KeyValuePair<int, Guid> appr in approvals.OrderByDescending(a => a.Key))
                {
                    Entity entApproval = new Entity("spectra_approval");
                    entApproval["spectra_name"] = appr.Key.ToString() + " approval";
                    entApproval["spectra_opportunityid"] = new EntityReference(context.PrimaryEntityName, context.PrimaryEntityId);
                    entApproval["spectra_orderby"] = new OptionSetValue(Convert.ToInt32(optionValuePreFix + appr.Key.ToString()));
                    //entApproval["spectra_approver"] = new EntityReference("systemuser", appr.Value);
                    entApproval["ownerid"] = new EntityReference("systemuser", appr.Value);
                    if (approvalId != Guid.Empty)
                        entApproval["spectra_nextapprovalid"] = new EntityReference("spectra_approval", approvalId);//new Guid("DF20F192-F7A0-E711-80E4-005056AD0689")); //
                    if (appr.Key == firstApproval)//if (appr.Key == 0)
                    {
                        entApproval["spectra_approvalrequesteddate"] = DateTime.Now;
                        entApproval["statecode"] = new OptionSetValue(0);
                        entApproval["statuscode"] = new OptionSetValue(111260000);
                    }
                    approvalId = service.Create(entApproval);
                    ApprovalCreated.Set(executionContext, true);

                    if (appr.Key == firstApproval)//if (appr.Key == 0)
                    {
                        //entCollOppProd = getOppProducts(service, context.PrimaryEntityId, true);
                        //byte[] pdfData = PdfGeneratorAndPostProcessor(service, context.PrimaryEntityId, entCollOppProd);

                        DOAHelper helper = new DOAHelper();

                        Entity entApprover = service.Retrieve("systemuser", appr.Value, new ColumnSet("fullname"));

                        string approver = (entApprover.Contains("fullname") ? entApprover["fullname"].ToString() : "Approver");
                        string emailbody = helper.getEmailBody(service, context.PrimaryEntityId.ToString(), approver);

                        Entity entEmail = new Entity("email");
                        entEmail["subject"] = "Pending for your approval";
                        entEmail["description"] = emailbody;

                        Entity entTo = new Entity("activityparty");
                        entTo["partyid"] = new EntityReference("systemuser", entApprover.Id);
                        Entity[] entToList = { entTo };
                        entEmail["to"] = entToList;

                        Entity Queue = helper.GetResultByAttribute(service, "queue", "name", "DOA Approval", "queueid");

                        if (Queue != null)
                        {
                            Entity entFrom = new Entity("activityparty");
                            //entFrom["partyid"] = new EntityReference("systemuser", new Guid("1A9B2FAD-7334-E711-80DE-000D3AF224B9"));// crm support user
                            entFrom["partyid"] = new EntityReference("queue", Queue.Id);
                            Entity[] entFromList = { entFrom };
                            entEmail["from"] = entFromList;
                        }
                        else
                            throw new InvalidPluginExecutionException("DOA approval not available");

                        Entity oppty = helper.GetResultByAttribute(service, "opportunity", "opportunityid", context.PrimaryEntityId.ToString(), "ownerid");

                        if (oppty != null)
                        {
                            Entity entcc = new Entity("activityparty");
                            //entFrom["partyid"] = new EntityReference("systemuser", new Guid("1A9B2FAD-7334-E711-80DE-000D3AF224B9"));// crm support user
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

                        //entEmail["regardingobjectid"] = new EntityReference("opportunity", context.PrimaryEntityId);
                        entEmail["regardingobjectid"] = new EntityReference("spectra_approval", approvalId);
                        Guid emailId = service.Create(entEmail);

                        //Entity attachment = new Entity("activitymimeattachment");
                        //attachment["subject"] = "Opportuntity Product approval details";
                        //attachment["filename"] = "OpportuntityProduct.pdf";
                        //attachment["body"] = Convert.ToBase64String(pdfData);
                        ////attachment["mimetype"] = "text/plain";
                        //attachment["attachmentnumber"] = 1;
                        //attachment["objectid"] = new EntityReference("email", emailId);
                        //attachment["objecttypecode"] = "email";
                        //service.Create(attachment);

                        //Send email
                        SendEmailRequest sendEmailReq = new SendEmailRequest()
                        {
                            EmailId = emailId,
                            IssueSend = true
                        };
                        SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);

                    }

                    //if (appr.Key == 0)
                    //{
                    //    //throw new InvalidPluginExecutionException("hi       " + approvalId.ToString());
                    //    SetStateRequest staReq = new SetStateRequest()
                    //    {
                    //        EntityMoniker = new EntityReference("spectra_approval", approvalId),
                    //        State = new OptionSetValue(0),
                    //        Status = new OptionSetValue(11126000)
                    //    };
                    //    //service.Execute(staReq);
                    //}
                    //isApprovalRequiredFlag = true;
                }

                Entity entOpp = new Entity("opportunity");
                entOpp.Id = context.PrimaryEntityId;
                entOpp["spectra_approvalrequiredflag"] = false;
                entOpp["statecode"] = new OptionSetValue(0);
                entOpp["statuscode"] = new OptionSetValue(569480013);
                service.Update(entOpp);
            }
        }

        public EntityCollection getOppProducts(IOrganizationService service, Guid oppId, Boolean retrieveOnlyApprovedRecords)
        {
            QueryExpression query = new QueryExpression("opportunityproduct");
            query.NoLock = true;
            query.ColumnSet = new ColumnSet("extendedamount", "priceperunit", "productid", "spectra_approvalrequried");
            query.Criteria.AddCondition(new ConditionExpression("opportunityid", ConditionOperator.Equal, oppId));
            if (retrieveOnlyApprovedRecords)
                query.Criteria.AddCondition(new ConditionExpression("spectra_approvalrequried", ConditionOperator.Equal, true));

            return service.RetrieveMultiple(query);
        }

        public EntityCollection getApprovalConfig(IOrganizationService service, string appConfigNameType, string prodseg)//, decimal percentAge)
        {
            QueryExpression query = new QueryExpression("spectra_approvalconfig");
            query.ColumnSet = new ColumnSet("spectra_approver", "spectra_name", "spectra_orderby", "spectra_minpercentage", "spectra_maxpercentage", "spectra_quantity");//spectra_percentage
            query.Criteria.AddCondition(new ConditionExpression("spectra_name", ConditionOperator.Equal, appConfigNameType.ToUpper()));

            if (prodseg != null)
                query.Criteria.AddCondition(new ConditionExpression("spectra_productsegment", ConditionOperator.Equal, prodseg));

            query.Orders.Add(new OrderExpression("spectra_minpercentage", OrderType.Ascending));
            query.Orders.Add(new OrderExpression("spectra_orderby", OrderType.Ascending));
            return service.RetrieveMultiple(query);
        }
    }
}
