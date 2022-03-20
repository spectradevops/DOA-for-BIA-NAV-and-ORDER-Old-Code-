using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Text.RegularExpressions;

namespace DOA
{
    public class UpdateOpportunityApprovalField : IPlugin
    {
        Int64 _ipAddNoApprovalLimit = 0;
        string searchData = "_IPADDRESS_";

        public UpdateOpportunityApprovalField(string unsecureConfig)
        {
            //load unsecureConfig If unsecureConfig is not null 
            if (!string.IsNullOrWhiteSpace(unsecureConfig))
                _ipAddNoApprovalLimit = Convert.ToInt64(unsecureConfig);
        }

        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = (IOrganizationService)factory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService trace = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            //if (context.Depth > 1)
            //    return;

            if (context.PrimaryEntityName.ToLower() == "opportunityproduct")
            {
                if (context.MessageName.ToLower() == "create")
                {
                    if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                    {
                        Entity entTraget = (Entity)context.InputParameters["Target"];
                        //if (entTraget.Contains("alletech_businesssegment") && ((EntityReference)entTraget["alletech_businesssegment"]).Name.ToLower() == "business")
                        {
                            //throw new InvalidPluginExecutionException(entTraget.Contains("productid").ToString() + "    " + entTraget.Contains("extendedamount").ToString() + "  " + ((Money)entTraget["extendedamount"]).Value + "      " + (((Money)entTraget["extendedamount"]).Value != 0).ToString());
                            if (entTraget.Contains("productid") && entTraget.Contains("extendedamount") && ((Money)entTraget["extendedamount"]).Value != 0)//&& ((EntityReference)entTraget["productid"]).Name.ToLower()
                            {
                                EntityReference prodId = (EntityReference)entTraget["productid"];

                                if (prodId.Name.ToLower().Contains("_ipaddress_"))
                                {
                                    EntityReference oppId = (EntityReference)entTraget["opportunityid"];
                                    Entity entOpp = service.Retrieve(oppId.LogicalName, oppId.Id, new ColumnSet("alletech_businesssegmentglb", "statuscode"));
                                    if (entOpp.Contains("alletech_businesssegmentglb") && ((EntityReference)entOpp["alletech_businesssegmentglb"]).Name.ToLower() == "business")
                                    {
                                        Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_businesssegmentlookup", "alletech_grossplaninvoicevalueinr"));
                                        if (entProd.Contains("alletech_businesssegmentlookup") && ((EntityReference)entProd["alletech_businesssegmentlookup"]).Name.ToLower() == "business")
                                        {
                                            //if (entProd.Contains("alletech_grossplaninvoicevalueinr"))
                                            {
                                                decimal extendedAmt = ((Money)entTraget["extendedamount"]).Value;

                                                var cnt = prodId.Name.Substring(prodId.Name.IndexOf(searchData) + searchData.Length);
                                                cnt = Regex.Replace(cnt, "[^0-9]+", string.Empty);
                                                Int64 count = Convert.ToInt64(cnt);

                                                decimal percentAge = extendedAmt / count;

                                                if (percentAge < _ipAddNoApprovalLimit)
                                                {
                                                    entOpp["spectra_approvalrequiredflag"] = true;
                                                    service.Update(entOpp);
                                                }
                                            }
                                        }
                                    }
                                }
                                else if (prodId.Name.ToLower().EndsWith("rc") || prodId.Name.ToLower().EndsWith("otc"))// || prodId.Name.ToLower().EndsWith("ip"))
                                {
                                    EntityReference oppId = (EntityReference)entTraget["opportunityid"];
                                    Entity entOpp = service.Retrieve(oppId.LogicalName, oppId.Id, new ColumnSet("alletech_businesssegmentglb", "statuscode"));
                                    if (entOpp.Contains("alletech_businesssegmentglb") && ((EntityReference)entOpp["alletech_businesssegmentglb"]).Name.ToLower() == "business")
                                    {
                                        Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_businesssegmentlookup", "alletech_grossplaninvoicevalueinr"));
                                        if (entProd.Contains("alletech_businesssegmentlookup") && ((EntityReference)entProd["alletech_businesssegmentlookup"]).Name.ToLower() == "business")
                                        {
                                            if (entProd.Contains("alletech_grossplaninvoicevalueinr"))
                                            {
                                                decimal extendedAmt = ((Money)entTraget["extendedamount"]).Value;
                                                decimal floorDisc = ((Money)entProd["alletech_grossplaninvoicevalueinr"]).Value;
                                                if (extendedAmt < floorDisc)
                                                {
                                                    entOpp["spectra_approvalrequiredflag"] = true;
                                                    service.Update(entOpp);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                else
                {
                    EntityCollection entCollOppProd = new EntityCollection();
                    Entity entOpp = new Entity();
                    bool isApprovalRequiredFlag = false;// true;
                    if (context.MessageName.ToLower() == "update")
                    {
                        if (context.PostEntityImages.Contains("PostImage") && context.PostEntityImages["PostImage"] is Entity)
                        {
                            trace.Trace("In update");

                            Entity entPost = context.PostEntityImages["PostImage"];
                            {
                                if (entPost.Contains("opportunityid"))
                                {
                                    entOpp.Id = ((EntityReference)entPost["opportunityid"]).Id;
                                }
                                if (entPost.Contains("productid"))
                                {
                                    EntityReference prodId = (EntityReference)entPost["productid"];

                                    EntityReference oppId = (EntityReference)entPost["opportunityid"];
                                    entOpp = service.Retrieve(oppId.LogicalName, oppId.Id, new ColumnSet("alletech_businesssegmentglb", "statuscode"));

                                    if (entOpp.Contains("alletech_businesssegmentglb") && ((EntityReference)entOpp["alletech_businesssegmentglb"]).Name.ToLower() == "business")
                                    {
                                        bool approval = false;

                                        Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_businesssegmentlookup", "alletech_plantype", "alletech_chargetype", "alletech_grossplaninvoicevalueinr"));

                                        decimal manualdiscount = 0;
                                        if (entPost.Attributes.Contains("manualdiscountamount"))
                                            manualdiscount = ((Money)entPost["manualdiscountamount"]).Value;

                                        decimal price = ((Money)entPost["priceperunit"]).Value;
                                        decimal extendedAmt = price - manualdiscount;
                                        int chargetype = entProd.GetAttributeValue<OptionSetValue>("alletech_chargetype").Value;
                                        int plantype = entProd.GetAttributeValue<OptionSetValue>("alletech_plantype").Value;

                                        //throw new InvalidPluginExecutionException("after variables");

                                        #region Ip Address
                                        if (plantype == 569480002 && chargetype == 569480001)
                                        {
                                            trace.Trace("If it is Add On");
                                            if (prodId.Name.ToLower().Contains("_ipaddress_"))
                                            {
                                                #region Commented on 17_March_2022 by Madhu
                                                var cnt = prodId.Name.Substring(prodId.Name.IndexOf(searchData) + searchData.Length);
                                                cnt = Regex.Replace(cnt, "[^0-9]+", string.Empty);
                                                Int64 count = Convert.ToInt64(cnt);
                                                #endregion
                                                //decimal percentAge = extendedAmt / count;

                                                if (extendedAmt < _ipAddNoApprovalLimit)
                                                {
                                                    trace.Trace("percentAge is less than _ipAddNoApprovalLimit");
                                                    approval = true;
                                                }
                                                else
                                                {
                                                    approval = false;
                                                }
                                            }
                                            #region New logic has been added on 17 March 2022
                                            else
                                            {
                                                if (extendedAmt < price)
                                                {
                                                    trace.Trace("if extendedAmt < floorDisc");
                                                    decimal percentAge = (price - extendedAmt) / price * 100;
                                                    percentAge = decimal.Round(percentAge, 2, MidpointRounding.AwayFromZero);

                                                    if (entPost.Contains("spectra_approvedpercentage") && percentAge <= (decimal)entPost["spectra_approvedpercentage"])
                                                    {
                                                        trace.Trace("percentAge <= (decimal)entPost[spectra_approvedpercentage]");
                                                        approval = false;
                                                    }
                                                    else if ((!entPost.Contains("spectra_approvedpercentage")) || (entPost.Contains("spectra_approvedpercentage") && percentAge > (decimal)entPost["spectra_approvedpercentage"]))
                                                    {
                                                        approval = true;
                                                    }
                                                }
                                                else
                                                    approval = false;
                                            }
                                            #endregion New logic has been added on 17 March 2022
                                        }
                                        #endregion

                                        #region RC|| OTC
                                        else if (plantype == 569480001&&(chargetype == 569480001 || chargetype == 569480002))
                                        {
                                            trace.Trace("if RC || OTC");
                                            //decimal floorDisc = ((Money)entProd["alletech_grossplaninvoicevalueinr"]).Value;
                                            if (extendedAmt < price)
                                            {
                                                trace.Trace("if extendedAmt < floorDisc");
                                                decimal percentAge = (price - extendedAmt) / price * 100;
                                                percentAge = decimal.Round(percentAge, 2, MidpointRounding.AwayFromZero);

                                                if (entPost.Contains("spectra_approvedpercentage") && percentAge <= (decimal)entPost["spectra_approvedpercentage"])
                                                {
                                                    trace.Trace("percentAge <= (decimal)entPost[spectra_approvedpercentage]");
                                                    approval = false;
                                                }
                                                else if ((!entPost.Contains("spectra_approvedpercentage")) || (entPost.Contains("spectra_approvedpercentage") && percentAge > (decimal)entPost["spectra_approvedpercentage"]))
                                                {
                                                    approval = true;
                                                }
                                            }
                                            else
                                                approval = false;
                                        }
                                        #endregion

                                        #region Approvals
                                        if (approval)
                                        {
                                            Entity entOppProdUpdate = new Entity(context.PrimaryEntityName);
                                            entOppProdUpdate.Id = context.PrimaryEntityId;
                                            entOppProdUpdate["spectra_approvalrequried"] = true;
                                            service.Update(entOppProdUpdate);
                                            entCollOppProd = getOppProducts(service, oppId.Id, null);
                                        }
                                        else
                                        {
                                            Entity entOppProdUpdate = new Entity(context.PrimaryEntityName);
                                            entOppProdUpdate.Id = context.PrimaryEntityId;
                                            entOppProdUpdate["spectra_approvalrequried"] = false;
                                            service.Update(entOppProdUpdate);
                                            entCollOppProd = getOppProducts(service, oppId.Id, context.PrimaryEntityId.ToString());
                                        }
                                        #endregion
                                    }
                                }
                            }
                        }
                    }
                    else if (context.MessageName.ToLower() == "delete")
                    {
                        trace.Trace("This plugin is fired for delete message");
                        if (context.PreEntityImages.Contains("PreImage") && context.PreEntityImages["PreImage"] is Entity)
                        {
                            trace.Trace("Context contains Pre image");

                            Entity entPre = context.PreEntityImages["PreImage"];
                            {
                                if (entPre.Contains("opportunityid"))
                                {
                                    entOpp.Id = ((EntityReference)entPre["opportunityid"]).Id;
                                }
                                if (entPre.Contains("productid"))
                                {

                                    trace.Trace("Preimage has a product");
                                    EntityReference prodId = (EntityReference)entPre["productid"];
                                    Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_plantype", "alletech_chargetype"));
                                    int chargetype = entProd.GetAttributeValue<OptionSetValue>("alletech_chargetype").Value;
                                    int plantype = entProd.GetAttributeValue<OptionSetValue>("alletech_plantype").Value;

                                    if (plantype == 569480002 || chargetype == 569480001 || chargetype == 569480002)//Ip address||RC||OTC
                                    {
                                        trace.Trace("Product is either ipaddress or rc or otc");
                                        EntityReference oppId = (EntityReference)entPre["opportunityid"];
                                        entOpp = service.Retrieve(oppId.LogicalName, oppId.Id, new ColumnSet("alletech_businesssegmentglb", "statuscode"));
                                        trace.Trace("checked whether the oppty is for business segment");
                                        if (entOpp.Contains("alletech_businesssegmentglb") && ((EntityReference)entOpp["alletech_businesssegmentglb"]).Name.ToLower() == "business")
                                        {
                                            trace.Trace("get oppty products");
                                            entCollOppProd = getOppProducts(service, oppId.Id, context.PrimaryEntityId.ToString());
                                            trace.Trace("total products " + entCollOppProd.Entities.Count.ToString());
                                        }
                                    }
                                }
                            }
                        }
                    }
                    
                    foreach (Entity entOppProd in entCollOppProd.Entities)
                    {
                        #region commented
                        /*if (entOppProd.Contains("productid"))
                        {
                            trace.Trace("if oppty product contains product");
                            EntityReference prodId = (EntityReference)entOppProd["productid"];
                            if (prodId.Name.ToLower().Contains("_ipaddress_"))
                            {
                                trace.Trace("if oppty product contains product and is of ipaddress");
                                Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_businesssegmentlookup", "alletech_grossplaninvoicevalueinr"));
                                trace.Trace("get floor price");
                                if (entProd.Contains("alletech_businesssegmentlookup") && ((EntityReference)entProd["alletech_businesssegmentlookup"]).Name.ToLower() == "business")
                                {
                                    trace.Trace("Product is for business");
                                    //if (entProd.Contains("alletech_grossplaninvoicevalueinr"))
                                    {
                                        decimal extendedAmt = ((Money)entOppProd["extendedamount"]).Value;
                                        trace.Trace("get the extended amount of oppty product");
                                        var cnt = prodId.Name.Substring(prodId.Name.IndexOf(searchData) + searchData.Length);
                                        cnt = Regex.Replace(cnt, "[^0-9]+", string.Empty);
                                        Int64 count = Convert.ToInt64(cnt);

                                        decimal percentAge = extendedAmt / count;

                                        if (percentAge < _ipAddNoApprovalLimit)
                                        {
                                            trace.Trace("value below the floor price then update the flag to true");
                                            entOpp["spectra_approvalrequiredflag"] = true;
                                            service.Update(entOpp);
                                            trace.Trace("Opportunity updated with ApprovalRequiredFlag as True.");
                                            isApprovalRequiredFlag = true;
                                            if (((OptionSetValue)entOpp["statuscode"]).Value == 569480014)
                                            {
                                                entOpp["statecode"] = new OptionSetValue(0);
                                                entOpp["statuscode"] = new OptionSetValue(1);
                                            }
                                            trace.Trace("Opportunity Status updated");
                                            break;
                                        }
                                    }
                                }
                            }
                            else if (prodId.Name.ToLower().EndsWith("rc") || prodId.Name.ToLower().EndsWith("otc"))
                            {
                                trace.Trace("this is for rc or otc");
                                Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_businesssegmentlookup", "alletech_grossplaninvoicevalueinr"));
                                if (entProd.Contains("alletech_businesssegmentlookup") && ((EntityReference)entProd["alletech_businesssegmentlookup"]).Name.ToLower() == "business")
                                {
                                    trace.Trace("Business segment product");
                                    if (entProd.Contains("alletech_grossplaninvoicevalueinr"))
                                    {
                                        trace.Trace("get the floor price");
                                        decimal extendedAmt = ((Money)entOppProd["extendedamount"]).Value;
                                        decimal floorDisc = ((Money)entProd["alletech_grossplaninvoicevalueinr"]).Value;
                                        if (extendedAmt < floorDisc)
                                        {
                                            trace.Trace("If the extended amt is less than floor then update the flag to true");
                                            //entOpp["spectra_approvalrequiredflag"] = true;
                                            //service.Update(entOpp);
                                            //isApprovalRequiredFlag = true;
                                            //if (((OptionSetValue)entOpp["statuscode"]).Value == 569480014)
                                            //{
                                            //    entOpp["statecode"] = new OptionSetValue(0);
                                            //    entOpp["statuscode"] = new OptionSetValue(1);
                                            //}
                                            //break;

                                            decimal percentAge = (floorDisc - extendedAmt) / floorDisc * 100;
                                            percentAge = decimal.Round(percentAge, 2, MidpointRounding.AwayFromZero);

                                            if (entOppProd.Contains("spectra_approvedpercentage") && percentAge <= (decimal)entOppProd["spectra_approvedpercentage"])
                                            {
                                                //Entity entOppProdUpdate = new Entity(context.PrimaryEntityName);
                                                //entOppProdUpdate.Id = context.PrimaryEntityId;
                                                //entOppProdUpdate["spectra_approvalrequried"] = false;
                                                //service.Update(entOppProdUpdate);
                                                entOpp["spectra_approvalrequiredflag"] = false;
                                                service.Update(entOpp);
                                            }
                                            else if ((!entOppProd.Contains("spectra_approvedpercentage")) || (entOppProd.Contains("spectra_approvedpercentage") && percentAge > (decimal)entOppProd["spectra_approvedpercentage"]))
                                            {
                                                //Entity entOppProdUpdate = new Entity(context.PrimaryEntityName);
                                                //entOppProdUpdate.Id = context.PrimaryEntityId;
                                                //entOppProdUpdate["spectra_approvalrequried"] = true;
                                                //service.Update(entOppProdUpdate);
                                                entOpp["spectra_approvalrequiredflag"] = true;
                                                service.Update(entOpp);
                                                isApprovalRequiredFlag = true;

                                                if (((OptionSetValue)entOpp["statuscode"]).Value == 569480014)
                                                {
                                                    entOpp["statecode"] = new OptionSetValue(0);
                                                    entOpp["statuscode"] = new OptionSetValue(1);
                                                }
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }*/
                        #endregion
                        //check for all products
                        if (entOppProd.Attributes.Contains("spectra_approvalrequried") && entOppProd.GetAttributeValue<bool>("spectra_approvalrequried"))
                        {
                            isApprovalRequiredFlag = true;
                            trace.Trace("approvalrequired is true");
                            break;
                        }
                    }
                    if (entOpp != new Entity() && entOpp != null && entOpp.Id != null)//&& entOpp != new Entity())
                    {
                        entOpp = service.Retrieve("opportunity", entOpp.Id, new ColumnSet("alletech_businesssegmentglb", "spectra_approvalrequiredflag"));
                        if (entOpp.Contains("alletech_businesssegmentglb") && ((EntityReference)entOpp["alletech_businesssegmentglb"]).Name.ToLower() == "business")
                        {
                            if (entOpp.Attributes.Contains("spectra_approvalrequiredflag"))
                            {
                                if (entOpp.GetAttributeValue<bool>("spectra_approvalrequiredflag") != isApprovalRequiredFlag)
                                {
                                    Entity oppty = new Entity(entOpp.LogicalName);
                                    oppty.Id = entOpp.Id;
                                    oppty["spectra_approvalrequiredflag"] = isApprovalRequiredFlag;
                                    service.Update(oppty);
                                    trace.Trace("flag updated");
                                }
                            }
                            else if (isApprovalRequiredFlag)
                            {
                                Entity oppty = new Entity(entOpp.LogicalName);
                                oppty.Id = entOpp.Id;
                                oppty["spectra_approvalrequiredflag"] = isApprovalRequiredFlag;
                                service.Update(oppty);
                                trace.Trace("flag updated");
                            }
                        }
                    }
                }

                #region old
                //else if (context.MessageName.ToLower() == "update")
                //{
                //    if (context.PostEntityImages.Contains("PostImage") && context.PostEntityImages["PostImage"] is Entity)
                //    {
                //        Entity entPost = context.PostEntityImages["PostImage"];
                //        //if (entTraget.Contains("alletech_businesssegment") && ((EntityReference)entTraget["alletech_businesssegment"]).Name.ToLower() == "business")
                //        {
                //            if (entPost.Contains("productid"))// && entTraget.Contains("extendedamount") && ((Money)entTraget["extendedamount"]).Value != 0.0m)//&& ((EntityReference)entTraget["productid"]).Name.ToLower()
                //            {
                //                EntityReference prodId = (EntityReference)entPost["productid"];
                //                if (prodId.Name.ToLower().EndsWith("rc") || prodId.Name.ToLower().EndsWith("otc") || prodId.Name.ToLower().EndsWith("ip"))
                //                {
                //                    EntityReference oppId = (EntityReference)entPost["opportunityid"];
                //                    Entity entOpp = service.Retrieve(oppId.LogicalName, oppId.Id, new ColumnSet("alletech_businesssegment"));
                //                    if (entOpp.Contains("alletech_businesssegment") && ((EntityReference)entOpp["alletech_businesssegment"]).Name.ToLower() == "business")
                //                    {
                //                        EntityCollection entCollOppProd = getOppProducts(service, oppId.Id, null);
                //                        foreach (Entity entOppProd in entCollOppProd.Entities)
                //                        {
                //                            //Entity entProd = srvice.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_businesssegmentlookup", "alletech_grossplaninvoicevalueinr"));
                //                            //if (entProd.Contains("alletech_businesssegmentlookup") && ((EntityReference)entProd["alletech_businesssegmentlookup"]).Name.ToLower() == "business")
                //                            {
                //                                if (entOppProd.Contains("alletech_grossplaninvoicevalueinr"))
                //                                {
                //                                    decimal extendedAmt = ((Money)entPost["extendedamount"]).Value;
                //                                    decimal floorDisc = ((Money)entOppProd["alletech_grossplaninvoicevalueinr"]).Value;
                //                                    if (extendedAmt <= floorDisc)
                //                                    {
                //                                        //Entity entOpp = new Entity("opportunity");
                //                                        //entOpp.Id = ((EntityReference)entTraget["opportunityid"]).Id;
                //                                        entOpp["spectra_approvalrequiredflag"] = true;
                //                                        service.Update(entOpp);
                //                                        break;
                //                                    }
                //                                }
                //                            }
                //                        }
                //                    }
                //                }
                //            }
                //        }
                //    }

                //    //    if (context.PostEntityImages.Contains("PostImage") && context.PostEntityImages["PostImage"] is Entity)
                //    //    {
                //    //        Entity entPost = (Entity)context.PostEntityImages["PostImage"];
                //    //        if (entPost.Contains("alletech_businesssegment") && ((EntityReference)entPost["alletech_businesssegment"]).Name.ToLower() == "business")
                //    //        {
                //    //            if (entPost.Contains("productid") && entPost.Contains("extendedamount") && ((Money)entPost["extendedamount"]).Value != 0.0m)//&& ((EntityReference)entTraget["productid"]).Name.ToLower()
                //    //            {
                //    //                EntityReference prodId = (EntityReference)entPost["productid"];
                //    //                if (prodId.Name.ToLower().EndsWith("rc") || prodId.Name.ToLower().EndsWith("otc") || prodId.Name.ToLower().EndsWith("ip"))
                //    //                {
                //    //                    Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_businesssegmentlookup", "alletech_grossplaninvoicevalueinr"));
                //    //                    if (entProd.Contains("alletech_businesssegmentlookup") && ((EntityReference)entProd["alletech_businesssegmentlookup"]).Name.ToLower() == "business")
                //    //                    {
                //    //                        if (entProd.Contains("alletech_grossplaninvoicevalueinr"))
                //    //                        {
                //    //                            decimal extendedAmt = ((Money)entPost["extendedamount"]).Value;
                //    //                            decimal floorDisc = ((Money)entProd["alletech_grossplaninvoicevalueinr"]).Value;
                //    //                            if (extendedAmt <= floorDisc)
                //    //                            {
                //    //                                Entity entOpp = new Entity("opportunity");
                //    //                                entOpp.Id = ((EntityReference)entPost["opportunityid"]).Id;
                //    //                                entOpp["spectra_approvalrequiredflag"] = true;
                //    //                                service.Update(entOpp);

                //    //                                //EntityCollection entCollAppConfig = getApprovalConfig(service, (prodId.Name.ToLower().EndsWith("rc") ? "RC" : (prodId.Name.ToLower().EndsWith("otc") ? "OTC" : "IP")));
                //    //                                //decimal percentAge = floorDisc / extendedAmt * 100;
                //    //                                //bool isApprovalRequiredFlag = false;
                //    //                                //foreach (Entity entAppConfig in entCollAppConfig.Entities)
                //    //                                //{
                //    //                                //    if (entAppConfig.Contains("spectra_percentage") && percentAge < (decimal)entAppConfig["spectra_percentage"])
                //    //                                //    {
                //    //                                //        Entity entApproval = new Entity("spectra_approval");
                //    //                                //        entApproval[""] = "";
                //    //                                //        service.Create(entApproval);
                //    //                                //        isApprovalRequiredFlag = true;
                //    //                                //    }
                //    //                                //}

                //    //                                //if (isApprovalRequiredFlag && entTraget.Contains("opportunityid"))
                //    //                                //{
                //    //                                //    Entity entOpp = new Entity("opportunity");
                //    //                                //    entOpp.Id = ((EntityReference)entTraget["opportunityid"]).Id;
                //    //                                //    entOpp["spectra_approvalrequiredflag"] = true;
                //    //                                //    service.Update(entOpp);
                //    //                                //}
                //    //                            }
                //    //                        }
                //    //                    }
                //    //                }
                //    //            }
                //    //        }
                //    //    }
                //    //}
                //    //{
                //    //    //if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                //    //    if (context.PostEntityImages.Contains("PostImage") && context.PostEntityImages["PostImage"] is Entity)
                //    //    {
                //    //        Entity entPost = (Entity)context.PostEntityImages["PostImage"];
                //    //        if (entPost.Contains("alletech_businesssegment") && ((EntityReference)entPost["alletech_businesssegment"]).Name.ToLower() == "business")
                //    //        {
                //    //            if (entPost.Contains("productid") && entPost.Contains("manualdiscountamount"))
                //    //            {
                //    //                EntityReference prodId = (EntityReference)entPost["productid"];
                //    //                Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_grossplaninvoicevalueinr"));//
                //    //                if (entProd.Contains("alletech_grossplaninvoicevalueinr"))
                //    //                {
                //    //                    decimal manualDisc = (decimal)entProd["manualdiscountamount"];
                //    //                    decimal floorDisc = (decimal)entProd["alletech_grossplaninvoicevalueinr"];
                //    //                    if (manualDisc > floorDisc)
                //    //                    {
                //    //                        if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                //    //                        {
                //    //                            Entity entTraget = (Entity)context.InputParameters["Target"];
                //    //                            entPost["spectra_approvalrequiredflag"] = true;
                //    //                        }
                //    //                    }
                //    //                }
                //    //            }
                //    //        }
                //    //    }
                //    //}
                //    else if (context.MessageName.ToLower() == "delete")
                //    {
                //        //if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                //        if (context.PreEntityImages.Contains("PreImage") && context.PreEntityImages["PreImage"] is Entity)
                //        {
                //            Entity entPost = (Entity)context.PreEntityImages["PreImage"];
                //            if (entPost.Contains("productid") && entPost.Contains("manualdiscountamount"))
                //            {
                //                EntityReference prodId = (EntityReference)entPost["productid"];
                //                Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_grossplaninvoicevalueinr"));//
                //                if (entProd.Contains("alletech_grossplaninvoicevalueinr"))
                //                {
                //                    decimal manualDisc = (decimal)entProd["manualdiscountamount"];
                //                    decimal floorDisc = (decimal)entProd["alletech_grossplaninvoicevalueinr"];
                //                    if (manualDisc > floorDisc)
                //                    {
                //                        if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                //                        {
                //                            Entity entTraget = (Entity)context.InputParameters["Target"];
                //                            entPost["spectra_approvalrequiredflag"] = true;
                //                        }
                //                    }
                //                }
                //            }
                //        }

                //        ////if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                //        //if (context.PostEntityImages.Contains("PostImage") && context.PostEntityImages["PostImage"] is Entity)
                //        //{
                //        //    Entity entPost = (Entity)context.PostEntityImages["PostImage"];
                //        //    if (entPost.Contains("productid") && entPost.Contains("manualdiscountamount"))
                //        //    {
                //        //        EntityReference prodId = (EntityReference)entPost["productid"];
                //        //        Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_grossplaninvoicevalueinr"));//
                //        //        if (entProd.Contains("alletech_grossplaninvoicevalueinr"))
                //        //        {
                //        //            decimal manualDisc = (decimal)entProd["manualdiscountamount"];
                //        //            decimal floorDisc = (decimal)entProd["alletech_grossplaninvoicevalueinr"];
                //        //            if (manualDisc > floorDisc)
                //        //            {
                //        //                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                //        //                {
                //        //                    Entity entTraget = (Entity)context.InputParameters["Target"];
                //        //                    entPost["spectra_approvalrequiredflag"] = true;
                //        //                }
                //        //            }
                //        //        }
                //        //    }
                //        //}
                //    }
                //}
                #endregion old

            }
        }

        public EntityCollection getOppProducts(IOrganizationService service, Guid oppId, string deleteOppProdID)
        {
            QueryExpression query = new QueryExpression("opportunityproduct");
            query.NoLock = true;
            query.ColumnSet = new ColumnSet("extendedamount", "productid", "spectra_approvedpercentage", "spectra_approvalrequried");
            query.Criteria.AddCondition(new ConditionExpression("opportunityid", ConditionOperator.Equal, oppId));
            if (deleteOppProdID != null && deleteOppProdID != "")
                query.Criteria.AddCondition(new ConditionExpression("opportunityproductid", ConditionOperator.NotEqual, deleteOppProdID));

            return service.RetrieveMultiple(query);
        }

    }
}
