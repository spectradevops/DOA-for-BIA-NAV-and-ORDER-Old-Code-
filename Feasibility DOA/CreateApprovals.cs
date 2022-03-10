using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Feasibility_DOA
{
    public class CreateApprovals : CodeActivity
    {
        [Input("OpportunityID")]
        [RequiredArgument]
        public InArgument<string> OpportunityID { get; set; }

        [Input("City")]
        [RequiredArgument]
        [ReferenceTarget("alletech_city")]
        public InArgument<EntityReference> City { get; set; }

        [Input("user")]
        [RequiredArgument]
        [ReferenceTarget("systemuser")]
        public InArgument<EntityReference> User { get; set; }

        //[Input("ClientURL")]
        //[RequiredArgument]
        //public InArgument<string> ClientURL { get; set; }

        //[Output("ApprovalCreated")]
        //public OutArgument<bool> ApprovalCreated { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService traceService = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(User.Get(executionContext).Id);

            if (context.PrimaryEntityName.ToLower() == "alletech_feasibility")
            {
                string BusinessSeg = string.Empty;
                string accountManager = string.Empty;
                string oppGUID = string.Empty;
                string productName = string.Empty;
                string feasibilityID = string.Empty;
                string billCycle = string.Empty;
                string oppID = OpportunityID.Get(executionContext);
                Dictionary<int, Guid> approvals = new Dictionary<int, Guid>();
                Entity feasibDetails = service.Retrieve("alletech_feasibility", context.PrimaryEntityId, new ColumnSet("alletech_subreason", "createdby", "alletech_feasibilityidd", "alletech_busiensssegment", "alletech_product", "alletech_opportunity"));
                traceService.Trace("got Feasibity details");
                if (feasibDetails.Attributes.Contains("alletech_subreason"))
                {
                    if (feasibDetails.GetAttributeValue<OptionSetValue>("alletech_subreason").Value == 111260000)
                    {
                        traceService.Trace("got alletech_subreason details");
                        string feasFetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                  <entity name='alletech_feasibility'>
                                                    <attribute name='alletech_feasibilityid' />
                                                    <attribute name='alletech_feasibilityidd' />
                                                    <attribute name='alletech_routetype' />
                                                    <attribute name='alletech_feasiblitystatus' />
                                                    <attribute name='alletech_opportunityid' />
                                                    <order attribute='alletech_routetype' descending='false' />
                                                    <filter type='and'>
                                                      <condition attribute='alletech_thirdpartyinstallation' operator='eq' value='1' />
                                                    </filter>
                                                    <link-entity name='opportunity' from='opportunityid' to='alletech_opportunity' alias='aj'>
                                                      <filter type='and'>
                                                        <condition attribute='alletech_oppurtunityid' operator='eq' value='" + oppID + @"' />
                                                      </filter>
                                                    </link-entity>
                                                  </entity>
                                                </fetch>";
                        EntityCollection feas_coll = service.RetrieveMultiple(new FetchExpression(feasFetch));
                        if (feas_coll.Entities.Count > 0)
                        {
                            traceService.Trace("Feasib Count: " + feas_coll.Entities.Count.ToString());
                            foreach (Entity FES in feas_coll.Entities)
                            {
                                if (FES.Attributes.Contains("alletech_feasiblitystatus"))
                                {
                                    if (((OptionSetValue)FES["alletech_feasiblitystatus"]).Value == 1 && FES.GetAttribute‌​‌​Value<bool>("alletech_routetype") == false)
                                    {
                                        traceService.Trace("alletech_feasiblitystatus: Primarly");
                                        string approvalFetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                      <entity name='spectra_approvalconfig'>
                                                        <attribute name='spectra_name' />
                                                        <attribute name='spectra_approver' />
                                                        <attribute name='spectra_approvalconfigid' />
                                                        <order attribute='spectra_name' descending='false' />
                                                        <filter type='and'>
                                                          <condition attribute='spectra_name' operator='eq' value='FEASIB_MSD' />
                                                        </filter>
                                                      </entity>
                                                    </fetch>";

                                        EntityCollection MSDcoll = service.RetrieveMultiple(new FetchExpression(approvalFetch));
                                        #region Adding MSD HEAD                                    

                                        if (MSDcoll.Entities.Count > 0)
                                        {
                                            traceService.Trace("MSDcoll.Entities.Count: " + MSDcoll.Entities.Count.ToString());
                                            if (MSDcoll.Entities[0].Contains("spectra_approver"))
                                                approvals.Add(0, MSDcoll.Entities[0].GetAttributeValue<EntityReference>("spectra_approver").Id);
                                        }
                                        else
                                            throw new InvalidPluginExecutionException("SRM Head not added in approval config");
                                        #endregion

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
                                            entApproval["spectra_feasibility"] = new EntityReference(context.PrimaryEntityName, context.PrimaryEntityId);
                                            entApproval["spectra_orderby"] = new OptionSetValue(Convert.ToInt32(optionValuePreFix + appr.Key.ToString()));
                                            entApproval["ownerid"] = new EntityReference("systemuser", appr.Value);
                                            if (approvalId != Guid.Empty)
                                                entApproval["spectra_nextapprovalid"] = new EntityReference("spectra_approval", approvalId);//new Guid("DF20F192-F7A0-E711-80E4-005056AD0689")); //
                                            if (appr.Key == firstApproval)//if (appr.Key == 0)
                                            {
                                                entApproval["spectra_approvalrequesteddate"] = DateTime.Now;
                                                entApproval["statecode"] = new OptionSetValue(0);
                                                entApproval["statuscode"] = new OptionSetValue(111260000);
                                            }
                                            traceService.Trace("Before Approval create");
                                            approvalId = service.Create(entApproval);
                                            traceService.Trace("after Approval create");
                                            // ApprovalCreated.Set(executionContext, true);

                                            if (appr.Key == firstApproval)//if (appr.Key == 0)
                                            {

                                                FeasibilityDOAHelper helper = new FeasibilityDOAHelper();
                                                Entity entApprover = service.Retrieve("systemuser", appr.Value, new ColumnSet("fullname"));
                                                string approver = (entApprover.Contains("fullname") ? entApprover["fullname"].ToString() : "Approver");


                                                #region Parameters                                              
                                                if (feasibDetails.Attributes.Contains("alletech_busiensssegment"))
                                                {
                                                    BusinessSeg = ((EntityReference)feasibDetails.Attributes["alletech_busiensssegment"]).Name.ToString();
                                                    traceService.Trace("BusinessSeg");
                                                }
                                                if (feasibDetails.Attributes.Contains("createdby"))
                                                {
                                                    accountManager = ((EntityReference)feasibDetails.Attributes["createdby"]).Name.ToString();
                                                    traceService.Trace("accountManager");
                                                }
                                                if (feasibDetails.Attributes.Contains("alletech_product"))
                                                {
                                                    productName = ((EntityReference)feasibDetails.Attributes["alletech_product"]).Name.ToString();
                                                    traceService.Trace("productName");
                                                    Entity produtColl = service.Retrieve("product", ((EntityReference)feasibDetails.Attributes["alletech_product"]).Id, new ColumnSet("alletech_billingcycle"));
                                                    if(produtColl.Attributes.Contains("alletech_billingcycle"))
                                                    {
                                                        billCycle = produtColl.GetAttributeValue<EntityReference>("alletech_billingcycle").Name.ToString();
                                                    }
                                                }
                                                if (feasibDetails.Attributes.Contains("alletech_feasibilityidd"))
                                                {
                                                    feasibilityID = feasibDetails.Attributes["alletech_feasibilityidd"].ToString();
                                                    traceService.Trace("feasibilityID");
                                                }
                                                if (feasibDetails.Attributes.Contains("alletech_opportunity"))
                                                {
                                                    oppGUID = ((EntityReference)feasibDetails.Attributes["alletech_opportunity"]).Id.ToString();
                                                    traceService.Trace("oppGUID: " + oppGUID);
                                                }
                                                #endregion
                                                traceService.Trace("Before Email body");
                                                string emailbody = helper.getEmailBody(service, oppGUID, approver, feasibilityID, productName, billCycle);
                                                traceService.Trace("After Email body");
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

                                                entEmail["regardingobjectid"] = new EntityReference("spectra_approval", approvalId);
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

                                        //Entity feas_Update = new Entity("alletech_feasibility");
                                        //feas_Update.Id = context.PrimaryEntityId;
                                        //feas_Update["spectra_approvalrequiredflag"] = new OptionSetValue(1);                                        
                                        //service.Update(feas_Update);
                                    }
                                }
                            }
                        }
                        #region unused Code
                        //  EntityReference prodseg = OpportunityDetails.GetAttributeValue<EntityReference>("alletech_productsegment");

                        // EntityCollection entCollOppProd = getOppProducts(service, context.PrimaryEntityId, false);

                        //foreach (Entity entOppProd in entCollOppProd.Entities)
                        //{
                        //    if (entOppProd.Contains("spectra_approvalrequried") && (bool)entOppProd["spectra_approvalrequried"])
                        //    {
                        //        if (entOppProd.Contains("productid"))
                        //        {
                        //            EntityReference prodId = (EntityReference)entOppProd["productid"];

                        //            Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_businesssegmentlookup", "alletech_plantype", "alletech_chargetype", "alletech_grossplaninvoicevalueinr"));
                        //            if (entProd.Contains("alletech_businesssegmentlookup") && ((EntityReference)entProd["alletech_businesssegmentlookup"]).Name.ToLower() == "business")
                        //            {
                        //                decimal percentAge = 0;
                        //                decimal extendedAmt = ((Money)entOppProd["extendedamount"]).Value;
                        //                int chargetype = entProd.GetAttributeValue<OptionSetValue>("alletech_chargetype").Value;
                        //                int plantype = entProd.GetAttributeValue<OptionSetValue>("alletech_plantype").Value;

                        //                if (plantype == 569480002 && chargetype == 569480001) //ip address
                        //                {
                        //                    var cnt = prodId.Name.Substring(prodId.Name.IndexOf(searchData) + searchData.Length);
                        //                    cnt = Regex.Replace(cnt, "[^0-9]+", string.Empty);
                        //                    Int64 count = Convert.ToInt64(cnt);
                        //                    decimal price = ((Money)entOppProd["priceperunit"]).Value;

                        //                    if (extendedAmt < 200)
                        //                    {
                        //                        percentAge = extendedAmt;
                        //                        EntityCollection entCollAppConfig = getApprovalConfig(service, "IPADDRESS", null);//, percentAge);
                        //                        foreach (Entity entAppConfig in entCollAppConfig.Entities)
                        //                        {
                        //                            if ((entAppConfig.Contains("spectra_quantity") && count >= Convert.ToInt64(entAppConfig["spectra_quantity"])) || !entAppConfig.Contains("spectra_quantity"))
                        //                            {
                        //                                if (entAppConfig.Contains("spectra_minpercentage") && entAppConfig.Contains("spectra_maxpercentage")
                        //                                    && (percentAge >= ((decimal)entAppConfig["spectra_minpercentage"])) && (percentAge <= ((decimal)entAppConfig["spectra_maxpercentage"])))
                        //                                {
                        //                                    if (entAppConfig.Contains("spectra_orderby") && entAppConfig.Contains("spectra_approver"))
                        //                                    {
                        //                                        KeyValuePair<int, Guid> apprDuplicate = approvals.FirstOrDefault(a => a.Key == Convert.ToInt16(entAppConfig.FormattedValues["spectra_orderby"]));
                        //                                        if (apprDuplicate.Equals(new KeyValuePair<int, Guid>()))
                        //                                        {
                        //                                            //approvals.Add(Convert.ToInt16(entAppConfig.FormattedValues["spectra_orderby"]), ((EntityReference)entAppConfig["spectra_approver"]).Id);
                        //                                            KeyValuePair<int, Guid> apprDuplicateCond2 = approvals.FirstOrDefault(a => a.Value == ((EntityReference)entAppConfig["spectra_approver"]).Id);
                        //                                            if (apprDuplicateCond2.Equals(new KeyValuePair<int, Guid>()))
                        //                                                approvals.Add(Convert.ToInt16(entAppConfig.FormattedValues["spectra_orderby"]), ((EntityReference)entAppConfig["spectra_approver"]).Id);
                        //                                            else
                        //                                            {
                        //                                                approvals.Remove(apprDuplicateCond2.Key);
                        //                                                approvals.Add(Convert.ToInt16(entAppConfig.FormattedValues["spectra_orderby"]), ((EntityReference)entAppConfig["spectra_approver"]).Id);
                        //                                            }
                        //                                        }
                        //                                    }
                        //                                }
                        //                            }
                        //                        }
                        //                    }
                        //                }

                        //                //RC || OTC
                        //                else if (chargetype == 569480001 || chargetype == 569480002)// 
                        //                {
                        //                    if (entProd.Contains("alletech_grossplaninvoicevalueinr"))
                        //                    {
                        //                        decimal floorDisc = ((Money)entProd["alletech_grossplaninvoicevalueinr"]).Value;
                        //                        if (extendedAmt < floorDisc)
                        //                        {
                        //                            percentAge = (floorDisc - extendedAmt) / floorDisc * 100;
                        //                            percentAge = decimal.Round(percentAge, 2, MidpointRounding.AwayFromZero);

                        //                            EntityCollection entCollAppConfig = getApprovalConfig(service, (chargetype == 569480001 ? "RC" : "OTC"), prodseg.Id.ToString());//, percentAge);
                        //                            foreach (Entity entAppConfig in entCollAppConfig.Entities)
                        //                            {
                        //                                if (entAppConfig.Contains("spectra_minpercentage") && entAppConfig.Contains("spectra_maxpercentage")
                        //                                    && (percentAge >= ((decimal)entAppConfig["spectra_minpercentage"])) && (percentAge <= ((decimal)entAppConfig["spectra_maxpercentage"])))
                        //                                {
                        //                                    if (entAppConfig.Contains("spectra_orderby") && entAppConfig.Contains("spectra_approver"))
                        //                                    {
                        //                                        KeyValuePair<int, Guid> apprDuplicate = approvals.FirstOrDefault(a => a.Key == Convert.ToInt16(entAppConfig.FormattedValues["spectra_orderby"]));
                        //                                        if (apprDuplicate.Equals(new KeyValuePair<int, Guid>()))
                        //                                        {
                        //                                            //approvals.Add(Convert.ToInt16(entAppConfig.FormattedValues["spectra_orderby"]), ((EntityReference)entAppConfig["spectra_approver"]).Id);
                        //                                            KeyValuePair<int, Guid> apprDuplicateCond2 = approvals.FirstOrDefault(a => a.Value == ((EntityReference)entAppConfig["spectra_approver"]).Id);
                        //                                            if (apprDuplicateCond2.Equals(new KeyValuePair<int, Guid>()))
                        //                                                approvals.Add(Convert.ToInt16(entAppConfig.FormattedValues["spectra_orderby"]), ((EntityReference)entAppConfig["spectra_approver"]).Id);
                        //                                            else
                        //                                            {
                        //                                                approvals.Remove(apprDuplicateCond2.Key);
                        //                                                approvals.Add(Convert.ToInt16(entAppConfig.FormattedValues["spectra_orderby"]), ((EntityReference)entAppConfig["spectra_approver"]).Id);
                        //                                            }
                        //                                        }
                        //                                    }
                        //                                }
                        //                            }
                        //                        }
                        //                    }
                        //                }
                        //            }
                        //        }
                        //    }
                        //}
                        #endregion
                    }
                }
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
