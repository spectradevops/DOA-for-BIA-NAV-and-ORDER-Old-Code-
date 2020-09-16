using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

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

            if (context.Depth > 1)
                return;

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
                                    Entity entOpp = service.Retrieve(oppId.LogicalName, oppId.Id, new ColumnSet("alletech_businesssegmentglb"));
                                    if (entOpp.Contains("alletech_businesssegmentglb") && ((EntityReference)entOpp["alletech_businesssegmentglb"]).Name.ToLower() == "business")
                                    {
                                        Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_businesssegmentlookup", "alletech_floorprice"));
                                        if (entProd.Contains("alletech_businesssegmentlookup") && ((EntityReference)entProd["alletech_businesssegmentlookup"]).Name.ToLower() == "business")
                                        {
                                            //if (entProd.Contains("alletech_floorprice"))
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
                                    Entity entOpp = service.Retrieve(oppId.LogicalName, oppId.Id, new ColumnSet("alletech_businesssegmentglb"));
                                    if (entOpp.Contains("alletech_businesssegmentglb") && ((EntityReference)entOpp["alletech_businesssegmentglb"]).Name.ToLower() == "business")
                                    {
                                        Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_businesssegmentlookup", "alletech_floorprice"));
                                        if (entProd.Contains("alletech_businesssegmentlookup") && ((EntityReference)entProd["alletech_businesssegmentlookup"]).Name.ToLower() == "business")
                                        {
                                            if (entProd.Contains("alletech_floorprice"))
                                            {
                                                decimal extendedAmt = ((Money)entTraget["extendedamount"]).Value;
                                                decimal floorDisc = ((Money)entProd["alletech_floorprice"]).Value;
                                                if (extendedAmt <= floorDisc)
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
                            Entity entPost = context.PostEntityImages["PostImage"];
                            {
                                if (entPost.Contains("productid"))
                                {
                                    EntityReference prodId = (EntityReference)entPost["productid"];
                                    if (prodId.Name.ToLower().Contains("_ipaddress_") || prodId.Name.ToLower().EndsWith("rc") || prodId.Name.ToLower().EndsWith("otc"))// || prodId.Name.ToLower().EndsWith("ip"))
                                    {
                                        EntityReference oppId = (EntityReference)entPost["opportunityid"];
                                        entOpp = service.Retrieve(oppId.LogicalName, oppId.Id, new ColumnSet("alletech_businesssegmentglb"));
                                        if (entOpp.Contains("alletech_businesssegmentglb") && ((EntityReference)entOpp["alletech_businesssegmentglb"]).Name.ToLower() == "business")
                                        {
                                            entCollOppProd = getOppProducts(service, oppId.Id, null);
                                        }
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
                                    //if (prodId.Name != "ENT_BIA10")
                                    //    throw new InvalidPluginExecutionException(prodId.Name);
                                    if (prodId.Name.ToLower().Contains("_ipaddress_") || prodId.Name.ToLower().EndsWith("rc") || prodId.Name.ToLower().EndsWith("otc"))// || prodId.Name.ToLower().EndsWith("ip"))
                                    {
                                        trace.Trace("Product is either ipaddress or rc or otc");
                                        EntityReference oppId = (EntityReference)entPre["opportunityid"];
                                        entOpp = service.Retrieve(oppId.LogicalName, oppId.Id, new ColumnSet("alletech_businesssegmentglb"));
                                        trace.Trace("checked whether the oppty is for business segment");
                                        if (entOpp.Contains("alletech_businesssegmentglb") && ((EntityReference)entOpp["alletech_businesssegmentglb"]).Name.ToLower() == "business")
                                        {
                                            trace.Trace("get oppty products");
                                            entCollOppProd = getOppProducts(service, oppId.Id, context.PrimaryEntityId.ToString());
                                            trace.Trace("total products "+entCollOppProd.Entities.Count.ToString());
                                        }
                                    }
                                }
                            }
                        }
                    }
                    //if (entCollOppProd.Entities.Count >= 0)
                    //    isApprovalRequiredFlag = false;

                    foreach (Entity entOppProd in entCollOppProd.Entities)
                    {
                        if (entOppProd.Contains("productid"))
                        {
                            trace.Trace("if oppty product contains product");
                            EntityReference prodId = (EntityReference)entOppProd["productid"];
                            if (prodId.Name.ToLower().Contains("_ipaddress_"))
                            {
                                trace.Trace("if oppty product contains product and is of ipaddress");
                                Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_businesssegmentlookup", "alletech_floorprice"));
                                trace.Trace("get floor price");
                                if (entProd.Contains("alletech_businesssegmentlookup") && ((EntityReference)entProd["alletech_businesssegmentlookup"]).Name.ToLower() == "business")
                                {
                                    trace.Trace("Product is for business");
                                    //if (entProd.Contains("alletech_floorprice"))
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
                                            isApprovalRequiredFlag = true;
                                            break;
                                        }
                                    }
                                }
                            }
                            else if (prodId.Name.ToLower().EndsWith("rc") || prodId.Name.ToLower().EndsWith("otc"))
                            {
                                trace.Trace("this is for rc or otc");
                                Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_businesssegmentlookup", "alletech_floorprice"));
                                if (entProd.Contains("alletech_businesssegmentlookup") && ((EntityReference)entProd["alletech_businesssegmentlookup"]).Name.ToLower() == "business")
                                {
                                    trace.Trace("Business segment product");
                                    if (entProd.Contains("alletech_floorprice"))
                                    {
                                        trace.Trace("get the floor price");
                                        decimal extendedAmt = ((Money)entOppProd["extendedamount"]).Value;
                                        decimal floorDisc = ((Money)entProd["alletech_floorprice"]).Value;
                                        if (extendedAmt <= floorDisc)
                                        {
                                            trace.Trace("If the extended amt is less than floor then update the flag to true");
                                            entOpp["spectra_approvalrequiredflag"] = true;
                                            service.Update(entOpp);
                                            isApprovalRequiredFlag = true;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (!isApprovalRequiredFlag && entOpp != null && entOpp.Id != null)//&& entOpp != new Entity())
                    {
                        trace.Trace("If the extended amt is greater than floor then update the flag to false"+entOpp.LogicalName);
                        entOpp["spectra_approvalrequiredflag"] = false;//isApprovalRequiredFlag;//
                        
                        entOpp.LogicalName = "opportunity";
                        
                        service.Update(entOpp);
                        trace.Trace("flag updated");
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
                //                            //Entity entProd = srvice.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_businesssegmentlookup", "alletech_floorprice"));
                //                            //if (entProd.Contains("alletech_businesssegmentlookup") && ((EntityReference)entProd["alletech_businesssegmentlookup"]).Name.ToLower() == "business")
                //                            {
                //                                if (entOppProd.Contains("alletech_floorprice"))
                //                                {
                //                                    decimal extendedAmt = ((Money)entPost["extendedamount"]).Value;
                //                                    decimal floorDisc = ((Money)entOppProd["alletech_floorprice"]).Value;
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
                //    //                    Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_businesssegmentlookup", "alletech_floorprice"));
                //    //                    if (entProd.Contains("alletech_businesssegmentlookup") && ((EntityReference)entProd["alletech_businesssegmentlookup"]).Name.ToLower() == "business")
                //    //                    {
                //    //                        if (entProd.Contains("alletech_floorprice"))
                //    //                        {
                //    //                            decimal extendedAmt = ((Money)entPost["extendedamount"]).Value;
                //    //                            decimal floorDisc = ((Money)entProd["alletech_floorprice"]).Value;
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
                //    //                Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_floorprice"));//
                //    //                if (entProd.Contains("alletech_floorprice"))
                //    //                {
                //    //                    decimal manualDisc = (decimal)entProd["manualdiscountamount"];
                //    //                    decimal floorDisc = (decimal)entProd["alletech_floorprice"];
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
                //                Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_floorprice"));//
                //                if (entProd.Contains("alletech_floorprice"))
                //                {
                //                    decimal manualDisc = (decimal)entProd["manualdiscountamount"];
                //                    decimal floorDisc = (decimal)entProd["alletech_floorprice"];
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
                //        //        Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_floorprice"));//
                //        //        if (entProd.Contains("alletech_floorprice"))
                //        //        {
                //        //            decimal manualDisc = (decimal)entProd["manualdiscountamount"];
                //        //            decimal floorDisc = (decimal)entProd["alletech_floorprice"];
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
            query.ColumnSet = new ColumnSet("extendedamount", "productid");
            query.Criteria.AddCondition(new ConditionExpression("opportunityid", ConditionOperator.Equal, oppId));
            if (deleteOppProdID != null && deleteOppProdID != "")
                query.Criteria.AddCondition(new ConditionExpression("opportunityproductid", ConditionOperator.NotEqual, deleteOppProdID));

            //LinkEntity lnPord = new LinkEntity("opportunityproduct", "product", "productid", "productid", JoinOperator.Inner);
            //lnPord.Columns.AddColumn("alletech_floorprice");
            //FilterExpression filPord = new FilterExpression(LogicalOperator.Or);
            //filPord.AddCondition("name", ConditionOperator.EndsWith, "RC");
            //filPord.AddCondition("name", ConditionOperator.EndsWith, "OTC");
            //filPord.AddCondition("name", ConditionOperator.EndsWith, "IP");
            //lnPord.LinkCriteria.AddFilter(filPord);

            //LinkEntity lnbusiSeg = new LinkEntity("product", "alletech_businesssegment", "alletech_businesssegmentlookup", "alletech_businesssegmentid", JoinOperator.Inner);
            //lnbusiSeg.LinkCriteria.AddCondition("alletech_name", ConditionOperator.Equal, "business");

            //lnPord.LinkEntities.Add(lnbusiSeg);
            //query.LinkEntities.Add(lnPord);

            return service.RetrieveMultiple(query);
        }

        public EntityCollection getApprovalConfig(IOrganizationService service, string appConfigNameType)
        {
            QueryExpression query = new QueryExpression("spectra_approvalconfig");
            query.ColumnSet = new ColumnSet("spectra_approver", "spectra_name", "spectra_orderby", "spectra_percentage");
            query.Criteria.AddCondition(new ConditionExpression("spectra_name", ConditionOperator.Equal, appConfigNameType.ToUpper()));
            query.Orders.Add(new OrderExpression("spectra_orderby", OrderType.Ascending));
            return service.RetrieveMultiple(query);
        }

    }
}
