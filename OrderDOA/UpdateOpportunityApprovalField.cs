using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Text.RegularExpressions;


namespace OrderDOA
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

            OrderDOAHelper helper = new OrderDOAHelper();

            if (context.Depth > 1)
                return;

            if (context.PrimaryEntityName.ToLower() == "salesorderdetail")
            {
                #region Create
                if (context.MessageName.ToLower() == "create")
                {
                    if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                    {
                        Entity entTraget = (Entity)context.InputParameters["Target"];
                        //if (entTraget.Contains("alletech_businesssegment") && ((EntityReference)entTraget["alletech_businesssegment"]).Name.ToLower() == "business")
                        {
                            //////throw new InvalidPluginExecutionException(entTraget.Contains("productid").ToString() + "    " + entTraget.Contains("extendedamount").ToString() + "  " + ((Money)entTraget["extendedamount"]).Value + "      " + (((Money)entTraget["extendedamount"]).Value != 0).ToString());
                            if (entTraget.Contains("productid") && entTraget.Contains("extendedamount") && ((Money)entTraget["extendedamount"]).Value != 0)//&& ((EntityReference)entTraget["productid"]).Name.ToLower()
                            {
                                EntityReference prodId = (EntityReference)entTraget["productid"];

                                if (prodId.Name.ToLower().Contains("_ipaddress_"))
                                {
                                    EntityReference oppId = (EntityReference)entTraget["salesorderid"];
                                    Entity entOpp = service.Retrieve(oppId.LogicalName, oppId.Id, new ColumnSet(true));

                                    Entity incident = service.Retrieve("incident", ((EntityReference)entOpp.Attributes["spectra_caseid"]).Id, new ColumnSet("alletech_businesssegment"));

                                    if (incident.Contains("alletech_businesssegment") && ((EntityReference)incident["alletech_businesssegment"]).Name.ToLower() == "business")
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
                                                    Entity order2 = new Entity(entOpp.LogicalName);
                                                    order2.Id = entOpp.Id;
                                                    order2["spectra_approvalrequiredflagorder"] = true;
                                                    service.Update(order2);
                                                }
                                            }
                                        }
                                    }
                                }
                                else if (prodId.Name.ToLower().EndsWith("rc"))// || prodId.Name.ToLower().EndsWith("otc"))// || prodId.Name.ToLower().EndsWith("ip"))
                                {
                                    EntityReference oppId = (EntityReference)entTraget["salesorderid"];
                                    Entity entOpp = service.Retrieve(oppId.LogicalName, oppId.Id, new ColumnSet(true));

                                    Entity incident = service.Retrieve("incident", ((EntityReference)entOpp.Attributes["spectra_caseid"]).Id, new ColumnSet("alletech_businesssegment"));

                                    if (incident.Contains("alletech_businesssegment") && ((EntityReference)incident["alletech_businesssegment"]).Name.ToLower() == "business")
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
                                                    Entity oppor = new Entity(oppId.LogicalName);
                                                    oppor.Id = oppId.Id;
                                                    oppor["spectra_approvalrequiredflagorder"] = true;
                                                    service.Update(oppor);

                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion

                #region Update Or Delete
                else
                {
                    EntityCollection entCollOppProd = new EntityCollection();
                    Entity Ord = new Entity();
                    Entity entPost = context.PostEntityImages["PostImage"];
                    Entity entPreimg = context.PreEntityImages["PostImage"];
                    Entity ProdEntity = null;
                    EntityReference Prodref = new EntityReference();
                    bool isApprovalRequiredFlag = false;// true;
                    string prodname = string.Empty;

                    #region Update
                    if (context.MessageName.ToLower() == "update")
                    {
                        //////throw new Exception("here");
                        trace.Trace("Inside Update");
                        if (context.PostEntityImages.Contains("PostImage") && context.PostEntityImages["PostImage"] is Entity)
                        {
                            trace.Trace("Contains Post Image");

                            if (entPost.Contains("salesorderid"))
                            {
                                trace.Trace("Contains salesorderid");
                                Ord.Id = ((EntityReference)entPost["salesorderid"]).Id;
                            }
                            if (entPost.Contains("productdescription"))
                            {
                                trace.Trace("Write in product");
                                QueryExpression query = new QueryExpression("product");
                                query.ColumnSet = new ColumnSet("name");
                                query.Criteria.AddCondition(new ConditionExpression("name", ConditionOperator.Equal, entPost.Attributes["productdescription"].ToString()));
                                EntityCollection prodcoll = service.RetrieveMultiple(query);
                                if (prodcoll != null && prodcoll.Entities.Count > 0)
                                {
                                    trace.Trace("Found product with Name : " + entPost.Attributes["productdescription"].ToString());
                                    ProdEntity = prodcoll.Entities[0];
                                    Prodref.Id = ProdEntity.Id;
                                    Prodref.LogicalName = ProdEntity.LogicalName;
                                    trace.Trace("Assignment Completed " + ProdEntity.Attributes["name"].ToString());
                                    Prodref.Name = ProdEntity.Attributes["name"].ToString();
                                    trace.Trace("Assignment Completed");
                                }
                            }
                            if (entPost.Contains("productid") || Prodref != null)
                            {
                                //trace.Trace("Contains productid");
                                EntityReference prodId = new EntityReference();
                                if (entPost.Contains("productid"))
                                {
                                    trace.Trace("Contains productid 1");
                                    prodId = (EntityReference)entPost["productid"];
                                }
                                else
                                {
                                    prodId = Prodref;
                                    trace.Trace("Contains Write-in Product " + prodId.Name.ToLower().Contains("otc"));
                                }

                                trace.Trace("Contains IP or RC or OTC");
                                EntityReference OrdId = (EntityReference)entPost["salesorderid"];
                                Ord = service.Retrieve(OrdId.LogicalName, OrdId.Id, new ColumnSet(true));

                                Entity incident = service.Retrieve("incident", ((EntityReference)Ord.Attributes["spectra_caseid"]).Id, new ColumnSet("alletech_businesssegment"));
                                Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_plantype", "alletech_chargetype"));

                                if (incident.Contains("alletech_businesssegment") && ((EntityReference)incident["alletech_businesssegment"]).Name.ToLower() == "business")
                                {
                                    trace.Trace("Contains Business Segment as Business");
                                    //decimal extendedAmt = ((Money)entPost["extendedamount"]).Value;
                                    decimal postmanualdiscount = 0, floorDisc =0, extendedAmt = 0;
                                    if (entPost.Attributes.Contains("manualdiscountamount"))                                    
                                        postmanualdiscount = ((Money)entPost["manualdiscountamount"]).Value;

                                    if (entPost.Attributes.Contains("priceperunit"))
                                        floorDisc = ((Money)entPost["priceperunit"]).Value;
                                                                       
                                    extendedAmt = floorDisc - postmanualdiscount;
                                    trace.Trace("extendedAmt");

                                    int chargetype = entProd.GetAttributeValue<OptionSetValue>("alletech_chargetype").Value;
                                    int plantype = entProd.GetAttributeValue<OptionSetValue>("alletech_plantype").Value;

                                    bool approval = false;

                                    #region IP
                                    if (plantype== 569480002 && chargetype == 569480001)
                                    {
                                        trace.Trace("If it is IP Address");
                                        var cnt = prodId.Name.Substring(prodId.Name.IndexOf(searchData) + searchData.Length);
                                        cnt = Regex.Replace(cnt, "[^0-9]+", string.Empty);
                                        Int64 count = Convert.ToInt64(cnt);

                                        decimal percentAge = extendedAmt / count;

                                        if (percentAge < _ipAddNoApprovalLimit)
                                        {
                                            trace.Trace("If percentAge < _ipAddNoApprovalLimit");
                                            approval = true;
                                        }
                                        else
                                        {
                                            trace.Trace("Else update spectra_approvalrequried to FLASE");
                                            approval = false;
                                        }
                                    }
                                    #endregion

                                    #region RC || OTC
                                    else if(plantype == 569480001 && chargetype == 569480001)// || chargetype == 569480002))
                                    {
                                        if (extendedAmt < floorDisc)
                                        {
                                            trace.Trace("if extendedAmt < floorDisc");
                                            decimal percentAge = (floorDisc - extendedAmt) / floorDisc * 100;
                                            percentAge = decimal.Round(percentAge, 2, MidpointRounding.AwayFromZero);

                                            if (entPost.Contains("spectra_approvedpercentage") && percentAge <= (decimal)entPost["spectra_approvedpercentage"])
                                            {
                                                trace.Trace("if entPost.Contains(spectra_approvedpercentage) && percentAge <= (decimal)entPost[spectra_approvedpercentage]");
                                                approval = false;

                                            }
                                            else if ((!entPost.Contains("spectra_approvedpercentage")) || (entPost.Contains("spectra_approvedpercentage") && percentAge > (decimal)entPost["spectra_approvedpercentage"]))
                                            {
                                                trace.Trace("if ((!entPost.Contains(spectra_approvedpercentage)) || (entPost.Contains(spectra_approvedpercentage) && percentAge > (decimal)entPost[spectra_approvedpercentage]))");
                                                approval = true;
                                            }
                                        }

                                    }
                                    #endregion

                                    #region approvals
                                    if (approval)
                                    {
                                        Entity entOppProdUpdate = new Entity(context.PrimaryEntityName);
                                        entOppProdUpdate.Id = context.PrimaryEntityId;
                                        entOppProdUpdate["spectra_approvalrequried"] = true;
                                        service.Update(entOppProdUpdate);
                                        trace.Trace("Updated spectra_approvalrequried to TRUE");

                                        entCollOppProd = getOppProducts(service, OrdId.Id, null);
                                    }
                                    else
                                    {
                                        Entity entOppProdUpdate = new Entity(context.PrimaryEntityName);
                                        entOppProdUpdate.Id = context.PrimaryEntityId;
                                        entOppProdUpdate["spectra_approvalrequried"] = false;
                                        service.Update(entOppProdUpdate);
                                        trace.Trace("Updated spectra_approvalrequried to false");

                                        entCollOppProd = getOppProducts(service, OrdId.Id, context.PrimaryEntityId.ToString());
                                    }
                                    #endregion
                                }
                            }
                        }
                    }
                    #endregion

                    #region Delete 
                    else if (context.MessageName.ToLower() == "delete")
                    {
                        trace.Trace("This plugin is fired for delete message");
                        if (context.PreEntityImages.Contains("PreImage") && context.PreEntityImages["PreImage"] is Entity)
                        {
                            trace.Trace("Context contains Pre image");

                            Entity entPre = context.PreEntityImages["PreImage"];
                            {
                                if (entPre.Contains("salesorderid"))
                                {
                                    Ord.Id = ((EntityReference)entPre["salesorderid"]).Id;
                                }
                                if (entPre.Contains("productid"))
                                {
                                    trace.Trace("Preimage has a product");
                                    EntityReference prodId = (EntityReference)entPre["productid"];
                                    //if (prodId.Name != "ENT_BIA10")
                                    //    ////throw new InvalidPluginExecutionException(prodId.Name);
                                    if (prodId.Name.ToLower().Contains("_ipaddress_") || prodId.Name.ToLower().EndsWith("rc") || prodId.Name.ToLower().EndsWith("otc"))// || prodId.Name.ToLower().EndsWith("ip"))
                                    {
                                        trace.Trace("Product is either ipaddress or rc or otc");
                                        EntityReference OrdId = (EntityReference)entPre["salesorderid"];
                                        Ord = service.Retrieve(OrdId.LogicalName, OrdId.Id, new ColumnSet(true));

                                        Entity incident = service.Retrieve("incident", ((EntityReference)Ord.Attributes["spectra_caseid"]).Id, new ColumnSet("alletech_businesssegment"));
                                        trace.Trace("checked whether the oppty is for business segment");
                                        if (incident.Contains("alletech_businesssegment") && ((EntityReference)incident["alletech_businesssegment"]).Name.ToLower() == "business")
                                        {
                                            trace.Trace("get oppty products");
                                            entCollOppProd = getOppProducts(service, OrdId.Id, context.PrimaryEntityId.ToString());
                                            trace.Trace("total products " + entCollOppProd.Entities.Count.ToString());
                                        }
                                    }
                                }
                            }
                        }
                    }
                    #endregion
                    //if (entCollOppProd.Entities.Count >= 0)
                    //    isApprovalRequiredFlag = false;

                    #region checking all remaining products

                    trace.Trace("Order products : " + entCollOppProd.Entities.Count);

                    foreach (Entity OrdProd in entCollOppProd.Entities)
                    {
                        //check for all products
                        if (OrdProd.Attributes.Contains("spectra_approvalrequried") && OrdProd.GetAttributeValue<bool>("spectra_approvalrequried"))
                        {
                            isApprovalRequiredFlag = true;
                            trace.Trace("approvalrequired is true");
                            break;
                        }
                    }
                    #endregion

                    #region updating order
                    ////throw new Exception("done"+ isApprovalRequiredFlag);
                    if (Ord != new Entity() && Ord != null && Ord.Id != null)//&& entOpp != new Entity())
                    {
                        Ord = service.Retrieve("salesorder", Ord.Id, new ColumnSet("spectra_approvalrequiredflagorder", "spectra_caseid"));
                        Entity incident = service.Retrieve("incident", ((EntityReference)Ord.Attributes["spectra_caseid"]).Id, new ColumnSet("alletech_businesssegment"));

                        if (incident.Contains("alletech_businesssegment") && ((EntityReference)incident["alletech_businesssegment"]).Name.ToLower() == "business")
                        {
                            trace.Trace("If the extended amt is greater than floor then update the flag to false" + Ord.LogicalName);

                            if (Ord.Attributes.Contains("spectra_approvalrequiredflagorder"))
                            {
                                if (Ord.GetAttributeValue<bool>("spectra_approvalrequiredflagorder") != isApprovalRequiredFlag)
                                {
                                    Entity order = new Entity(Ord.LogicalName);
                                    order.Id = Ord.Id;
                                    order["spectra_approvalrequiredflagorder"] = isApprovalRequiredFlag;
                                    service.Update(order);
                                    trace.Trace("flag updated");
                                }
                            }
                            else if (isApprovalRequiredFlag)
                            {
                                Entity order = new Entity(Ord.LogicalName);
                                order.Id = Ord.Id;
                                order["spectra_approvalrequiredflagorder"] = isApprovalRequiredFlag;
                                service.Update(order);
                                trace.Trace("flag updated");
                            }
                        }
                    }
                    #endregion
                }
                #endregion

                //throw new InvalidPluginExecutionException("eror");
            }
        }

        public EntityCollection getOppProducts(IOrganizationService service, Guid oppId, string deleteOppProdID)
        {
            QueryExpression query = new QueryExpression("salesorderdetail");
            query.ColumnSet = new ColumnSet("extendedamount", "spectra_approvalrequried", "productid", "spectra_approvedpercentage", "productdescription", "priceperunit", "manualdiscountamount");
            query.Criteria.AddCondition(new ConditionExpression("salesorderid", ConditionOperator.Equal, oppId));
            if (deleteOppProdID != null && deleteOppProdID != "")
                query.Criteria.AddCondition(new ConditionExpression("salesorderdetailid", ConditionOperator.NotEqual, deleteOppProdID));

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
