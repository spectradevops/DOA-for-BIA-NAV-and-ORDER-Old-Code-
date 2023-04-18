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

namespace Feasibility_DOA
{
    public class CreateApprovals : CodeActivity
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
                #region Email Parameters 24 Nov 2022
                string URL = string.Empty, Authkey = string.Empty, Action = string.Empty, from1 = string.Empty, to = string.Empty, cc = string.Empty, subject = string.Empty, content = string.Empty;
                URL = ML_URL.Get(executionContext);
                Authkey = ML_Authkey.Get(executionContext);
                Action = ML_Action.Get(executionContext);
                from1 = ML_from.Get(executionContext);
                #endregion
                string BusinessSeg = string.Empty;
                string accountManager = string.Empty;
                string oppGUID = string.Empty;
                string productName = string.Empty;
                string feasibilityID = string.Empty;
                string billCycle = string.Empty;
                string oppID = OpportunityID.Get(executionContext);
                string remarks = string.Empty;
                Dictionary<int, Guid> approvals = new Dictionary<int, Guid>();
                Entity feasibDetails = service.Retrieve("alletech_feasibility", context.PrimaryEntityId, new ColumnSet("alletech_remark", "alletech_subreason", "createdby", "alletech_feasibilityidd", "alletech_busiensssegment", "alletech_product", "alletech_opportunity"));
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
                                                    <attribute name='alletech_opportunity' />
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
                                                          <condition attribute='spectra_name' operator='eq' value='MBIAT_FEASIBILITY' />
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
                                            string apprvalFetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                                  <entity name='spectra_approval'>
                                                                    <attribute name='spectra_approvalid' />
                                                                    <attribute name='spectra_name' />
                                                                    <attribute name='createdon' />
                                                                    <order attribute='spectra_name' descending='false' />
                                                                    <filter type='and'>
                                                                      <condition attribute='spectra_feasibility' operator='eq' value='" + context.PrimaryEntityId + @"' />
                                                                    </filter>
                                                                  </entity>
                                                                </fetch>";
                                            EntityCollection apprvalColle = service.RetrieveMultiple(new FetchExpression(apprvalFetch));
                                            if (apprvalColle.Entities.Count == 0)
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
                                                        if (produtColl.Attributes.Contains("alletech_billingcycle"))
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
                                                    if (feasibDetails.Attributes.Contains("alletech_remark"))
                                                    {
                                                        remarks = feasibDetails.Attributes["alletech_remark"].ToString();
                                                        traceService.Trace("remarks: " + remarks);
                                                    }
                                                    #endregion
                                                   


                                                    #region Old Code 24 Nov 2022
                                                    //Entity entEmail = new Entity("email");
                                                    //entEmail["subject"] = "Pending for your approval #" + approvalId.ToString().ToUpper() + "#";
                                                    //entEmail["description"] = emailbody;

                                                    //Entity entTo = new Entity("activityparty");
                                                    //entTo["partyid"] = new EntityReference("systemuser", entApprover.Id);
                                                    //Entity[] entToList = { entTo };
                                                    //entEmail["to"] = entToList;

                                                    //Entity Queue = helper.GetResultByAttribute(service, "queue", "name", "DOA Approval", "queueid");

                                                    //if (Queue != null)
                                                    //{
                                                    //    Entity entFrom = new Entity("activityparty");
                                                    //    //entFrom["partyid"] = new EntityReference("systemuser", new Guid("1A9B2FAD-7334-E711-80DE-000D3AF224B9"));// crm support user
                                                    //    entFrom["partyid"] = new EntityReference("queue", Queue.Id);
                                                    //    Entity[] entFromList = { entFrom };
                                                    //    entEmail["from"] = entFromList;
                                                    //}
                                                    //else
                                                    //    throw new InvalidPluginExecutionException("DOA approval not available");


                                                    //Entity oppty = helper.GetResultByAttribute(service, "opportunity", "opportunityid", ((EntityReference)FES.Attributes["alletech_opportunity"]).Id.ToString(), "ownerid");
                                                    //if (oppty != null)
                                                    //{
                                                    //    Entity entcc1 = new Entity("activityparty");
                                                    //    entcc1["partyid"] = oppty.GetAttributeValue<EntityReference>("ownerid");
                                                    //    Entity[] entccList = { entcc1 };
                                                    //    entEmail["cc"] = entccList;
                                                    //}
                                                    //entEmail["regardingobjectid"] = new EntityReference("spectra_approval", approvalId);

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
                                                    subject = "Pending for your approval #" + approvalId.ToString().ToUpper() + "#";
                                                    traceService.Trace("Before Email body");
                                                    string emailbody = helper.getEmailBody(service, oppGUID, approver, feasibilityID, productName, billCycle, remarks, subject);
                                                    traceService.Trace("After Email body");
                                                    content = "Hi " + approver + ",\n" + emailbody.ToString();

                                                    Entity entApprover1 = service.Retrieve("systemuser", entApprover.Id, new ColumnSet("internalemailaddress"));

                                                    if (entApprover1.Attributes.Contains("internalemailaddress"))
                                                        to = entApprover1.Attributes["internalemailaddress"].ToString();

                                                    Entity opportunity = helper.GetResultByAttribute(service, "opportunity", "opportunityid", ((EntityReference)FES.Attributes["alletech_opportunity"]).Id.ToString(), "ownerid");                                                   
                                                    Entity opportunityOwner = service.Retrieve("systemuser", opportunity.GetAttributeValue<EntityReference>("ownerid").Id, new ColumnSet("internalemailaddress", "parentsystemuserid"));

                                                    if (opportunityOwner.Attributes.Contains("internalemailaddress"))
                                                        cc = opportunityOwner.Attributes["internalemailaddress"].ToString();                                                    

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
                                                                  <condition attribute='alletech_entityguid' operator='eq' value='" + approvalId + @"' />
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
                                                                log["alletech_name"] = "Feasibility Approval Email Pushed to ML: " + approvalId.ToString();
                                                                log["alletech_integrationwith"] = new OptionSetValue(4);
                                                                log["alletech_regardingentity"] = "spectra_approval";
                                                                log["alletech_entityguid"] = approvalId.ToString();
                                                                log["alletech_request"] = json.ToString();
                                                                log["alletech_responce"] = result.ToString();
                                                                service.Create(log);
                                                            }
                                                            else
                                                            {
                                                                Entity log = new Entity("alletech_integrationlog");
                                                                log["alletech_name"] = "Feasibility Approval Email Pushed to ML: " + approvalId.ToString();
                                                                log["alletech_integrationwith"] = new OptionSetValue(4);
                                                                log["alletech_regardingentity"] = "spectra_approval";
                                                                log["alletech_entityguid"] = approvalId.ToString();
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
    public class SendingEmail
    {
        public string Authkey { get; set; }
        public string Action { get; set; }
        public string from { get; set; }
        public string to { get; set; }
        public string cc { get; set; }
        public string subject { get; set; }
        public string content { get; set; }

    }
}
