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
    public class CreateApprovals : CodeActivity
    {
        [Input("user")]
        [RequiredArgument]
        [ReferenceTarget("systemuser")]
        public InArgument<EntityReference> User { get; set; }

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
            ITracingService trace = executionContext.GetExtension<ITracingService>();

            // throw new Exception("user id : "+context.UserId+ " user.get id "+ User.Get(executionContext).Id);

            if (context.PrimaryEntityName.ToLower() == "salesorder")
            {
                try
                {
                    int approvalCount = 0;
                    Dictionary<int, Guid> approvals = new Dictionary<int, Guid>();
                    //Entity account = null;
                    Entity ProdEntity = null;
                    EntityReference Prodref = new EntityReference();
                    OrderDOAHelper helper = new OrderDOAHelper();
                    string prodname = string.Empty;

                    Entity OrdDetails = service.Retrieve("salesorder", context.PrimaryEntityId, new ColumnSet("prioritycode", "customerid"));

                    traceService.Trace("got order details");
                    //throw new InvalidPluginExecutionException("got order details");

                    traceService.Trace("priority code : " + OrdDetails.GetAttributeValue<OptionSetValue>("prioritycode").Value);

                    //Permanent
                    if (OrdDetails.GetAttributeValue<OptionSetValue>("prioritycode").Value == 111260000)
                    {
                        #region Adding City head to approval list
                        //updated for City Head Change
                        /*if (OrdDetails.Attributes.Contains("customerid"))
                        {
                            account = service.Retrieve("account", ((EntityReference)OrdDetails.Attributes["customerid"]).Id, new ColumnSet("ownerid", "spectra_servicerelationshipmanagerid"));
                        }
                        //throw new Exception(""+context.PrimaryEntityId+" "+context.PrimaryEntityName);
                        if (account.Attributes.Contains("spectra_servicerelationshipmanagerid"))
                        {
                            traceService.Trace("Inside IF");
                            Entity oppowner = service.Retrieve("systemuser", ((EntityReference)account.Attributes["ownerid"]).Id, new ColumnSet("parentsystemuserid"));

                            if (oppowner.Attributes.Contains("parentsystemuserid"))
                            {
                                Entity entCity = service.Retrieve("alletech_city", ((EntityReference)oppowner.Attributes["alletech_operatingcity"]).Id, new ColumnSet("spectra_cityheadid"));
                                if (entCity.Contains("spectra_cityheadid"))
                                {
                                    approvals.Add(0, ((EntityReference)entCity["spectra_cityheadid"]).Id);
                                }
                                else
                                    throw new InvalidPluginExecutionException("City Head for the selected city is null, please contact System Administrator");
                            }
                            else
                            {
                                traceService.Trace("Inside else");
                                //throw new Exception("else part");
                                if (City.Get(executionContext).Id != null)
                                {
                                    Entity entCity = service.Retrieve("alletech_city", City.Get(executionContext).Id, new ColumnSet("spectra_cityheadid"));
                                    if (entCity.Contains("spectra_cityheadid"))
                                    {
                                        approvals.Add(0, ((EntityReference)entCity["spectra_cityheadid"]).Id);
                                    }
                                    else
                                        throw new InvalidPluginExecutionException("City Head for the selected city is null, please contact System Administrator");
                                }
                            }
                        }*/
                        #endregion

                        #region Adding SRM HEAD commented on 15-APRIL-2022
                        //EntityCollection SRMcoll = helper.getApprovalConfig(service, "SRMHead", "B2BUP_");

                        //if (SRMcoll.Entities.Count > 0)
                        //{
                        //    if (SRMcoll.Entities[0].Contains("spectra_approver"))
                        //        approvals.Add(0, SRMcoll.Entities[0].GetAttributeValue<EntityReference>("spectra_approver").Id);
                        //}
                        //else
                        //    throw new InvalidPluginExecutionException("SRM Head not added in approval config");
                        #endregion

                        #region Order Products level data capture

                        EntityCollection entCollOppProd = getOppProducts(service, context.PrimaryEntityId, false);
                        bool RCdiscount = false;
                        string produName = string.Empty;
                        foreach (Entity entOppProd in entCollOppProd.Entities)
                        {
                            if (entOppProd.Contains("spectra_approvalrequried") && (bool)entOppProd["spectra_approvalrequried"])
                            {
                                if (entOppProd.Contains("productdescription"))
                                {
                                    produName = entOppProd.Attributes["productdescription"].ToString();
                                    trace.Trace("Write in product");
                                    QueryExpression query = new QueryExpression("product");
                                    query.ColumnSet = new ColumnSet("name");
                                    query.Criteria.AddCondition(new ConditionExpression("name", ConditionOperator.Equal, entOppProd.Attributes["productdescription"].ToString()));
                                    EntityCollection prodcoll = service.RetrieveMultiple(query);
                                    if (prodcoll != null && prodcoll.Entities.Count > 0)
                                    {
                                        trace.Trace("Found product with Name : " + entOppProd.Attributes["productdescription"].ToString());
                                        ProdEntity = prodcoll.Entities[0];
                                        Prodref.Id = ProdEntity.Id;
                                        Prodref.LogicalName = ProdEntity.LogicalName;
                                        trace.Trace("Assignment Completed " + ProdEntity.Attributes["name"].ToString());
                                        Prodref.Name = ProdEntity.Attributes["name"].ToString();
                                        trace.Trace("Assignment Completed");
                                    }
                                }

                                if (entOppProd.Contains("productid") || Prodref != null)
                                {
                                    EntityReference prodId = new EntityReference();
                                    if (entOppProd.Contains("productid"))
                                    {
                                        trace.Trace("Contains productid");
                                        prodId = (EntityReference)entOppProd["productid"];
                                    }
                                    else
                                    {
                                        prodId = Prodref;
                                        trace.Trace("Contains Write-in Product " + prodId.Name.ToLower().Contains("otc"));
                                    }

                                    Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_businesssegmentlookup", "alletech_plantype", "alletech_chargetype", "alletech_grossplaninvoicevalueinr"));
                                    if (entProd.Contains("alletech_businesssegmentlookup") && ((EntityReference)entProd["alletech_businesssegmentlookup"]).Name.ToLower() == "business")
                                    {
                                        decimal percentAge = 0;
                                        decimal extendedAmt = ((Money)entOppProd["extendedamount"]).Value;
                                        int chargetype = entProd.GetAttributeValue<OptionSetValue>("alletech_chargetype").Value;
                                        int plantype = entProd.GetAttributeValue<OptionSetValue>("alletech_plantype").Value;

                                        #region Ip Address
                                        if (plantype == 569480002 && chargetype == 569480001)
                                        {
                                            var cnt = prodId.Name.Substring(prodId.Name.IndexOf(searchData) + searchData.Length);
                                            cnt = Regex.Replace(cnt, "[^0-9]+", string.Empty);
                                            Int64 count = Convert.ToInt64(cnt);

                                            percentAge = extendedAmt / count;

                                            EntityCollection entCollAppConfig = helper.getApprovalConfig(service, "IPADDRESS", null);//, percentAge);
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
                                        #endregion

                                        #region RC || OTC
                                        else if ((chargetype == 569480001 || produName.Contains("OTC")) && !RCdiscount) // chargetype == 569480002) && !RCdiscount)
                                        {
                                            if (entOppProd.Contains("priceperunit"))
                                            {
                                                decimal floorDisc = ((Money)entOppProd["priceperunit"]).Value;
                                                if (extendedAmt < floorDisc)
                                                {
                                                    percentAge = (floorDisc - extendedAmt) / floorDisc * 100;
                                                    percentAge = decimal.Round(percentAge, 2, MidpointRounding.AwayFromZero);

                                                    //EntityCollection entCollAppConfig = getApprovalConfig(service, (prodId.Name.ToLower().EndsWith("rc") ? "RC" : (prodId.Name.ToLower().EndsWith("otc") ? "OTC" : "IPADDRESS")));//, percentAge);
                                                    EntityCollection entCollAppConfig = helper.getApprovalConfigB2B(service, (chargetype == 569480001 ? "B2B_RC" : "B2B_OTC"), "");
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
                                                                    RCdiscount = true;
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
                                                    if (chargetype == 569480001 && approvals.Count > 1)
                                                    {
                                                        approvalCount = approvals.Count;
                                                        continue;
                                                    }
                                                }
                                            }
                                        }
                                        #endregion
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                    //temporary upgrade
                    else if (OrdDetails.GetAttributeValue<OptionSetValue>("prioritycode").Value == 111260001)
                    {
                        string fetchxml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                              <entity name='incident'>
                                                <attribute name='incidentid' />
                                                <attribute name='alletech_disposition' />
                                                <order attribute='title' descending='false' />
                                                <link-entity name='salesorder' from='spectra_caseid' to='incidentid' alias='aa'>
                                                  <filter type='and'>
                                                    <condition attribute='salesorderid' operator='eq' uiname='' uitype='salesorder' value='{" + context.PrimaryEntityId + @"}' />
                                                  </filter>
                                                </link-entity>
                                              </entity>
                                            </fetch>";

                        EntityCollection caseentity = service.RetrieveMultiple(new FetchExpression(fetchxml));

                        if (caseentity.Entities.Count > 0 && caseentity.Entities[0].Attributes.Contains("alletech_disposition"))
                        {
                            string subtype = caseentity.Entities[0].GetAttributeValue<EntityReference>("alletech_disposition").Name;

                            Entity approver = helper.GetResultByAttribute(service, "spectra_approvalconfig", "spectra_name", "B2BUP_" + subtype, "spectra_approver");

                            approvals.Add(0, ((EntityReference)approver["spectra_approver"]).Id);
                        }
                    }

                    //Home upgrade
                    else if (OrdDetails.GetAttributeValue<OptionSetValue>("prioritycode").Value == 111260002)
                    {
                        Entity approver = helper.GetResultByAttribute(service, "spectra_approvalconfig", "spectra_name", "B2CUP_Downgrade", "spectra_approver");
                        approvals.Add(0, ((EntityReference)approver["spectra_approver"]).Id);
                        traceService.Trace("In home upgrade");
                        //throw new InvalidPluginExecutionException("In home upgrade");
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
                    string order = null;
                    foreach (KeyValuePair<int, Guid> appr in approvals)
                    {
                        order += appr.Key.ToString() + " : ";
                        order += appr.Value.ToString() + " \n ";
                    }

                    //throw new InvalidPluginExecutionException("approvals count : " + approvals.Count);

                    foreach (KeyValuePair<int, Guid> appr in approvals.OrderByDescending(a => a.Key))
                    {
                        #region creating approval record
                        Entity entApproval = new Entity("spectra_approval");
                        entApproval["spectra_name"] = appr.Key.ToString() + " approval";
                        entApproval["spectra_orderid"] = new EntityReference(context.PrimaryEntityName, context.PrimaryEntityId);
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
                        #endregion

                        traceService.Trace("Created approval record");

                        //throw new InvalidPluginExecutionException("after approval");

                        #region if first response
                        if (appr.Key == firstApproval)
                        {
                            #region Assigning variables for body html
                            Entity accountdetails = null;
                            Entity casedetails = null;
                            string customersegment = string.Empty;
                            string billcycle = string.Empty;
                            string arc = string.Empty;

                            Entity entApprover = service.Retrieve("systemuser", appr.Value, new ColumnSet("fullname"));
                            Entity orderdetails = service.Retrieve("salesorder", context.PrimaryEntityId, new ColumnSet(true));

                            if (orderdetails.Attributes.Contains("customerid"))
                            {
                                Guid accid = ((EntityReference)(orderdetails.Attributes["customerid"])).Id;
                                accountdetails = service.Retrieve("account", accid, new ColumnSet(true));

                                if (accountdetails.Attributes.Contains("name"))
                                {
                                    Entity opp = helper.GetResultByAttribute(service, "opportunity", "alletech_companynamebusiness", accountdetails.Attributes["name"].ToString(), "spectra_customersegmentcode");

                                    if (opp != null && opp.Attributes.Contains("spectra_customersegmentcode"))
                                    {
                                        customersegment = opp.FormattedValues["spectra_customersegmentcode"];
                                    }
                                }

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

                            if (orderdetails.Attributes.Contains("spectra_caseid"))
                            {
                                Guid caseid = ((EntityReference)(orderdetails.Attributes["spectra_caseid"])).Id;
                                casedetails = service.Retrieve("incident", caseid, new ColumnSet(true));
                            }
                            if (orderdetails.Attributes.Contains("spectra_arc"))
                            {
                                arc = orderdetails.Attributes["spectra_arc"].ToString();
                            }

                            string approver = (entApprover.Contains("fullname") ? entApprover["fullname"].ToString() : "Approver");
                            #endregion

                            string emailbody = helper.getEmailBody(service, approver, orderdetails, accountdetails, casedetails, customersegment, billcycle, arc, traceService);

                            #region Creating EMail
                            Entity entEmail = new Entity("email");
                            entEmail["subject"] = "Pending for your approval";
                            entEmail["description"] = emailbody;

                            Entity Temp = new Entity();

                            Temp = new Entity("activityparty");
                            Temp["partyid"] = new EntityReference("systemuser", entApprover.Id);
                            Entity[] entToList = { Temp };
                            entEmail["to"] = entToList;

                            #region CC

                            List<Entity> entccList = new List<Entity>();
                            if (accountdetails.Attributes.Contains("spectra_servicerelationshipmanagerid"))
                            {
                                Temp = new Entity("activityparty");
                                Temp["partyid"] = accountdetails.GetAttributeValue<EntityReference>("spectra_servicerelationshipmanagerid");

                                entccList.Add(Temp);
                            }
                            if (approvalCount == 4)
                            {
                                EntityCollection CCcoll = helper.getApprovalConfig(service, "CC", "B2B_");
                                if (CCcoll.Entities.Count > 0)
                                {
                                    if (CCcoll.Entities[0].Contains("spectra_approver"))
                                    {
                                        Temp = new Entity("activityparty");
                                        Temp["partyid"] = CCcoll.Entities[0].GetAttributeValue<EntityReference>("spectra_approver");

                                        entccList.Add(Temp);
                                    }
                                }
                            }
                            Temp = new Entity("activityparty");
                            Temp["partyid"] = accountdetails.GetAttributeValue<EntityReference>("ownerid");
                            entccList.Add(Temp);

                            entEmail["cc"] = entccList.ToArray();
                            // }
                            #endregion

                            Entity Queue = helper.GetResultByAttribute(service, "queue", "name", "DOA Approval", "queueid");
                            if (Queue != null)
                            {
                                Temp = new Entity("activityparty");
                                Temp["partyid"] = new EntityReference("queue", Queue.Id);
                                Entity[] entFromList = { Temp };
                                entEmail["from"] = entFromList;
                            }
                            else
                                throw new InvalidPluginExecutionException("DOA approval not available");

                            entEmail["regardingobjectid"] = new EntityReference("spectra_approval", approvalId);
                            Guid emailId = service.Create(entEmail);
                            #endregion

                            //Send email
                            //throw new Exception("1");
                            SendEmailRequest sendEmailReq = new SendEmailRequest()
                            {
                                EmailId = emailId,
                                IssueSend = true
                            };
                            SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);
                        }
                        #endregion
                    }

                    #region updating order to waiting for approval
                    Entity entOpp = new Entity("salesorder");
                    entOpp.Id = context.PrimaryEntityId;
                    entOpp["spectra_approvalrequiredflagorder"] = false;
                    entOpp["statecode"] = new OptionSetValue(0);
                    entOpp["statuscode"] = new OptionSetValue(111260007);//waiting for approval
                    service.Update(entOpp);
                    #endregion

                }
                catch (Exception ex)
                {
                    throw new Exception("Error from Create Approvals" + ex.Message);
                }
            }
        }

        public EntityCollection getOppProducts(IOrganizationService service, Guid oppId, Boolean retrieveOnlyApprovedRecords)
        {
            QueryExpression query = new QueryExpression("salesorderdetail");
            query.NoLock = true;
            query.ColumnSet = new ColumnSet("extendedamount", "productid", "priceperunit", "spectra_approvalrequried", "productdescription");
            query.Criteria.AddCondition(new ConditionExpression("salesorderid", ConditionOperator.Equal, oppId));
            if (retrieveOnlyApprovedRecords)
                query.Criteria.AddCondition(new ConditionExpression("spectra_approvalrequried", ConditionOperator.Equal, true));

            query.Orders.Add(new OrderExpression("productid", OrderType.Descending));

            return service.RetrieveMultiple(query);
        }
    }
}
