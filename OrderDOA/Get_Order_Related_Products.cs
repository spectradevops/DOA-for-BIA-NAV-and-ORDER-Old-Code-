using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrderDOA
{
    public class Get_Order_Related_Products : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            EntityCollection pricelistcollection = null;
            Guid Normal_useProduct_Id = Guid.Empty;
            int total_norml_prd_count = 0;
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));

            IOrganizationService service = factory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            if (context.Depth > 1)
            {
                return;
            }
            try
            {
                tracingService.Trace("Plugin Execution Started");
                Entity _order = context.PostEntityImages["PostImage"];
                string subtypename = string.Empty;
                Money newOTC = new Money();
                //Entity _order = (Entity)service.Retrieve("salesorder", context.PrimaryEntityId, new ColumnSet("spectra_product", "spectra_getrelatedproducts", "customerid", "spectra_caseid"));
                Entity _account = (Entity)service.Retrieve("account", ((EntityReference)(_order.Attributes["customerid"])).Id, new ColumnSet("alletech_businesssegment", "alletech_buildingname", "alletech_product"));
                Entity _case = (Entity)service.Retrieve("incident", ((EntityReference)(_order.Attributes["spectra_caseid"])).Id, new ColumnSet("alletech_disposition"));
                EntityReference subtype = (EntityReference)_case.Attributes["alletech_disposition"];
                subtypename = subtype.Name;
                tracingService.Trace("Order,Account and Post Image Values Retrieved");
                if (_order.Contains("spectra_getrelatedproducts"))
                {
                    if (_order.Attributes["spectra_getrelatedproducts"].ToString() == "Yes")
                    {
                        tracingService.Trace("GetRelatedProduct Value is Yes");
                        if (_account.Attributes.Contains("alletech_businesssegment"))
                        {
                            string segment = ((EntityReference)_account.Attributes["alletech_businesssegment"]).Name;

                            if (segment == "Business")
                            {
                                tracingService.Trace("Business Segment Value is Business");

                                if (_order.Contains("alletech_productpkgcount"))
                                {
                                    total_norml_prd_count = Convert.ToInt32(_order["alletech_productpkgcount"]);
                                    if (total_norml_prd_count < 0)
                                        total_norml_prd_count = 0;
                                }
                                else
                                    total_norml_prd_count = 0;

                                if (_account.Attributes.Contains("alletech_buildingname"))
                                {
                                    tracingService.Trace("Account is having Building Value");

                                    Entity AccountPorduct = null;
                                    Entity OrderPorduct = null;
                                    int AccountBandwidth = 0;
                                    int OrderBandwidth = 0;

                                    //Contains Product on account
                                    if (_account.Attributes.Contains("alletech_product"))
                                    {
                                        tracingService.Trace("Account is having Product");
                                        AccountPorduct = service.Retrieve("product", ((EntityReference)_account.Attributes["alletech_product"]).Id, new ColumnSet(true));
                                        if (AccountPorduct.Attributes.Contains("alletech_bandwidthmaster"))
                                        {
                                            AccountBandwidth = ((OptionSetValue)AccountPorduct.Attributes["alletech_bandwidthmaster"]).Value;
                                            tracingService.Trace("AccountPorduct is having Bandwdith : " + AccountBandwidth);
                                        }
                                        else
                                        {
                                            throw new Exception("Bandwidth is empty for the current Product, please map the bandwidth on product" + ((EntityReference)_account.Attributes["alletech_product"]).Id);
                                        }
                                    }
                                    else
                                    {
                                        throw new Exception("Account is not associated with any of the product");
                                    }
                                    //Contains Product on Order
                                    if (_order.Attributes.Contains("spectra_product"))
                                    {
                                        tracingService.Trace("Order is having Product");
                                        OrderPorduct = service.Retrieve("product", ((EntityReference)_order.Attributes["spectra_product"]).Id, new ColumnSet(true));
                                        if (OrderPorduct.Attributes.Contains("alletech_bandwidthmaster"))
                                        {
                                            OrderBandwidth = ((OptionSetValue)OrderPorduct.Attributes["alletech_bandwidthmaster"]).Value;
                                            tracingService.Trace("Order Porduct is having Bandwdith : " + OrderBandwidth);
                                        }
                                        else
                                        {
                                            throw new Exception("Bandwidth is empty for the Upgraded Product, please map the bandwidth on product");
                                        }
                                    }

                                    #region checking config
                                    if (subtypename != string.Empty && subtypename != "Downgrade")
                                    {
                                        #region checking Opportunity products for existing products
                                        /*EntityCollection Oppor = Getdetails(service, tracingService, "opportunity", "parentaccountid", _account.Id);

                                        if (Oppor != null && Oppor.Entities.Count > 0)
                                        {
                                            EntityCollection OpporProduct = Getdetails(service, tracingService, "opportunityproduct", "opportunityid", Oppor.Entities[0].Id);
                                            foreach (Entity OP in OpporProduct.Entities)
                                            {
                                                if (((((EntityReference)OP.Attributes["productid"])).Name).Contains("OTC"))
                                                {
                                                    tracingService.Trace("Product Name contains OTC");
                                                    Money otcmoney = (Money)OP.Attributes["priceperunit"];

                                                    tracingService.Trace("OrderBandwidth : " + OrderBandwidth);
                                                    tracingService.Trace("AccountBandwidth : " + AccountBandwidth);

                                                    if (OrderBandwidth == AccountBandwidth)
                                                    {
                                                        newOTC.Value = 0;
                                                    }
                                                    else
                                                    {
                                                        //retrieve OTC from Config Entity
                                                        EntityCollection OTCConfig = GetOTCdetails(service, tracingService, "spectra_onetimecharges", new OptionSetValue(111260000), new Money(otcmoney.Value), new OptionSetValue(OrderBandwidth), new OptionSetValue(AccountBandwidth));

                                                        if (OTCConfig != null && OTCConfig.Entities.Count > 0)
                                                        {
                                                            //if Previous OTC on Opportunity is 25000
                                                            if (otcmoney.Value == 25000)
                                                            {
                                                                tracingService.Trace("Previous OTC Value is 25000");
                                                                if (OTCConfig.Entities[0].Attributes.Contains("spectra_newotc"))
                                                                    newOTC = (Money)OTCConfig.Entities[0].Attributes["spectra_newotc"];
                                                                else
                                                                    throw new Exception("New OTC value is not configured in the OTC config");
                                                            }
                                                            //if Previous OTC on Opportunity is 100000
                                                            else if (otcmoney.Value == 100000)
                                                            {
                                                                tracingService.Trace("Previous OTC Value is 100000");
                                                                if (OTCConfig.Entities[0].Attributes.Contains("spectra_newotc"))
                                                                    newOTC = (Money)OTCConfig.Entities[0].Attributes["spectra_newotc"];
                                                                else
                                                                    throw new Exception("New OTC value is not configured in the OTC config");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            throw new Exception("OTC Config doesnt exist for required upgrade");
                                                        }
                                                    }
                                                }
                                            }
                                        }*/
                                        #endregion

                                        #region Checking product relation and getting products
                                        //else
                                        //{
                                        tracingService.Trace("Checking product relation and getting products");
                                        QueryExpression query1_Child = new QueryExpression("alletech_productparent_productchild");
                                        query1_Child.ColumnSet = new ColumnSet(true);
                                        ConditionExpression cond1_Child = new ConditionExpression("productidone", ConditionOperator.Equal, ((EntityReference)_account.Attributes["alletech_product"]).Id);
                                        query1_Child.Criteria.AddCondition(cond1_Child);
                                        EntityCollection coll1_Child = service.RetrieveMultiple(query1_Child);

                                        tracingService.Trace("child count : " + coll1_Child.Entities.Count);

                                        if (coll1_Child != null && coll1_Child.Entities.Count > 0)
                                        {
                                            foreach (Entity product2 in coll1_Child.Entities)
                                            {
                                                Entity childProd = service.Retrieve("product", (Guid)product2["productidtwo"], new ColumnSet(true));

                                                if ((childProd.Attributes["name"].ToString()).Contains("OTC"))
                                                {
                                                    Money otcmoney = (Money)childProd.Attributes["alletech_grossplaninvoicevalueinr"];
                                                    //retrieve OTC from Config Entity
                                                    EntityCollection OTCConfig = GetOTCdetails(service, tracingService, "spectra_onetimecharges", new OptionSetValue(111260000), new Money(otcmoney.Value), new OptionSetValue(OrderBandwidth), new OptionSetValue(AccountBandwidth));

                                                    if (OTCConfig != null && OTCConfig.Entities.Count > 0)
                                                    {
                                                        //if Previous OTC on Opportunity is 25000
                                                        if (otcmoney.Value == 25000)
                                                        {
                                                            tracingService.Trace("Previous OTC Value is 25000");
                                                            if (OTCConfig.Entities[0].Attributes.Contains("spectra_newotc"))
                                                                newOTC = (Money)OTCConfig.Entities[0].Attributes["spectra_newotc"];
                                                            else
                                                                throw new Exception("New OTC value is not configured in the OTC config");

                                                        }
                                                        //if Previous OTC on Opportunity is 100000
                                                        else if (otcmoney.Value == 100000)
                                                        {
                                                            tracingService.Trace("Previous OTC Value is 100000");
                                                            if (OTCConfig.Entities[0].Attributes.Contains("spectra_newotc"))
                                                                newOTC = (Money)OTCConfig.Entities[0].Attributes["spectra_newotc"];
                                                            else
                                                                throw new Exception("New OTC value is not configured in the OTC config");
                                                        }
                                                    }
                                                    else
                                                    {
                                                        throw new Exception("OTC Config doesnt exist for required upgrade");
                                                    }
                                                }
                                            }
                                        }
                                        //}
                                        #endregion
                                    }
                                    #endregion

                                    EntityReference building = (EntityReference)_account.Attributes["alletech_buildingname"];

                                    QueryExpression query = new QueryExpression("pricelevel");
                                    query.ColumnSet = new ColumnSet(true);
                                    query.Criteria = new FilterExpression();
                                    query.Criteria.AddCondition("alletech_building", ConditionOperator.Equal, building.Id);

                                    pricelistcollection = service.RetrieveMultiple(query);

                                    tracingService.Trace("Building Pricelist value retrieved");

                                    if (pricelistcollection != null)
                                    {
                                        if (pricelistcollection.Entities.Count > 0)
                                        {
                                            Entity pricelist = pricelistcollection.Entities[0];

                                            //Get Order Products Associated to Orders
                                            QueryExpression queryorderproduct = new QueryExpression("salesorderdetail");
                                            queryorderproduct.ColumnSet = new ColumnSet(true);

                                            ConditionExpression cond = new ConditionExpression("salesorderid", ConditionOperator.Equal, _order.Id);
                                            queryorderproduct.Criteria.AddCondition(cond);
                                            EntityCollection OrderProductsCollection = service.RetrieveMultiple(queryorderproduct);
                                            tracingService.Trace("Order Products records retrieved");

                                            if (OrderProductsCollection.Entities.Count > 0)
                                            {
                                                foreach (Entity orderproduct in OrderProductsCollection.Entities)
                                                {
                                                    tracingService.Trace("Order Products Contains records");
                                                    Entity _useProduct = service.Retrieve("product", ((EntityReference)orderproduct.Attributes["productid"]).Id, new ColumnSet(true));

                                                    if (_useProduct.Attributes.Contains("alletech_plantype") && _useProduct.Attributes.Contains("alletech_chargetype"))
                                                    {
                                                        OptionSetValue plan = (OptionSetValue)_useProduct.Attributes["alletech_plantype"];
                                                        OptionSetValue chargetype = (OptionSetValue)_useProduct.Attributes["alletech_chargetype"];


                                                        if ((plan.Value == 569480001 && chargetype.Value == 569480000) || (plan.Value == 569480000 && chargetype.Value == 569480000))//(normal and package)or(retention and package)
                                                        {
                                                            Normal_useProduct_Id = _useProduct.Id;
                                                            ++total_norml_prd_count;
                                                        }
                                                    }

                                                    QueryExpression query1 = new QueryExpression("alletech_productparent_productchild");
                                                    query1.ColumnSet = new ColumnSet("productidtwo");
                                                    ConditionExpression cond1 = new ConditionExpression("productidone", ConditionOperator.Equal, _useProduct.Id);
                                                    query1.Criteria.AddCondition(cond1);
                                                    EntityCollection childproducts = service.RetrieveMultiple(query1);

                                                    tracingService.Trace("Child Products retrieval Completed");

                                                    if (childproducts.Entities.Count > 0)
                                                    {
                                                        foreach (Entity item1 in childproducts.Entities)
                                                        {
                                                            tracingService.Trace("Product is having child products");
                                                            Guid childpro = new Guid((item1.Attributes["productidtwo"].ToString()));
                                                            Entity _new_producttest = service.Retrieve("product", childpro, new ColumnSet(true));

                                                            if (subtypename != string.Empty && subtypename == "Downgrade")
                                                            {
                                                                #region For Degrade

                                                                //For Downgrade Remove OTC
                                                                //if (!(_new_producttest.Attributes["name"].ToString()).Contains("OTC")) //  Commented on 26th April 2022
                                                                // {
                                                                OptionSetValue statuscode = (OptionSetValue)_new_producttest.Attributes["statuscode"];

                                                                if (statuscode.Value == 1)
                                                                {
                                                                    QueryExpression query2 = new QueryExpression("productpricelevel");
                                                                    query2.ColumnSet = new ColumnSet(true);
                                                                    FilterExpression filter = new FilterExpression(LogicalOperator.And);
                                                                    ConditionExpression cond2 = new ConditionExpression("pricelevelid", ConditionOperator.Equal, pricelist.Id);
                                                                    ConditionExpression cond3 = new ConditionExpression("productid", ConditionOperator.Equal, _new_producttest.Id);
                                                                    filter.AddCondition(cond2);
                                                                    filter.AddCondition(cond3);
                                                                    query2.Criteria.AddFilter(filter);
                                                                    EntityCollection coll2 = service.RetrieveMultiple(query2);
                                                                    tracingService.Trace("Price list item retrieval completed" + coll2.Entities.Count);
                                                                    if (coll2.Entities.Count > 0)
                                                                    {
                                                                        foreach (Entity item3 in coll2.Entities)
                                                                        {

                                                                            Entity obj_listItem = service.Retrieve(item3.LogicalName, item3.Id, new ColumnSet(true));
                                                                            tracingService.Trace("Before Productid");
                                                                            Guid productidorder = ((EntityReference)(obj_listItem.Attributes["productid"])).Id;
                                                                            tracingService.Trace("Before amount");
                                                                            Money amount = (Money)obj_listItem.Attributes["amount"];
                                                                            tracingService.Trace("Before uomid");
                                                                            EntityReference uom = (EntityReference)obj_listItem.Attributes["uomid"];
                                                                            tracingService.Trace("Assignment Completed");

                                                                            //Create Order Product Records
                                                                            Entity new_Opp_Product = new Entity("salesorderdetail");

                                                                            new_Opp_Product["productid"] = new EntityReference("product", productidorder);
                                                                            new_Opp_Product["priceperunit"] = amount;
                                                                            new_Opp_Product["salesorderid"] = new EntityReference(_order.LogicalName, _order.Id);
                                                                            new_Opp_Product["uomid"] = uom;
                                                                            new_Opp_Product["quantity"] = (decimal)1;
                                                                            service.Create(new_Opp_Product);
                                                                            tracingService.Trace("Order Product create for child Product and Price List");

                                                                        }

                                                                        #region Check For Childs For Related Product
                                                                        if (_new_producttest.Attributes.Contains("alletech_chargetype"))
                                                                        {
                                                                            if (((OptionSetValue)(_new_producttest.Attributes["alletech_chargetype"])).Value == 569480000)
                                                                            {
                                                                                tracingService.Trace("If Charge Type is Package then retrieve child product related childs");
                                                                                QueryExpression query1_Child = new QueryExpression("alletech_productparent_productchild");
                                                                                query1_Child.ColumnSet = new ColumnSet(true);

                                                                                ConditionExpression cond1_Child = new ConditionExpression("productidone", ConditionOperator.Equal, _new_producttest.Id);

                                                                                query1_Child.Criteria.AddCondition(cond1_Child);
                                                                                EntityCollection coll1_Child = service.RetrieveMultiple(query1_Child);

                                                                                if (childproducts.Entities.Count > 0)
                                                                                {
                                                                                    foreach (Entity item1_Child in coll1_Child.Entities)
                                                                                    {
                                                                                        tracingService.Trace("Child product is having Child products");
                                                                                        Guid childchildpro = new Guid((item1.Attributes["productidtwo"].ToString()));
                                                                                        Entity _new_producttest_Child = service.Retrieve("product", childchildpro, new ColumnSet(true));


                                                                                        QueryExpression query2_Child = new QueryExpression("productpricelevel");
                                                                                        query2_Child.ColumnSet = new ColumnSet(true);
                                                                                        FilterExpression filter_Child = new FilterExpression(LogicalOperator.And);
                                                                                        ConditionExpression cond2_Child = new ConditionExpression("pricelevelid", ConditionOperator.Equal, pricelist.Id);
                                                                                        ConditionExpression cond3_Child = new ConditionExpression("productid", ConditionOperator.Equal, _new_producttest_Child.Id);

                                                                                        filter_Child.AddCondition(cond2_Child);
                                                                                        filter_Child.AddCondition(cond3_Child);
                                                                                        query2_Child.Criteria.AddFilter(filter_Child);
                                                                                        EntityCollection coll2_Child = service.RetrieveMultiple(query2_Child);
                                                                                        tracingService.Trace("Child-child product related procelist item retrieval completed");
                                                                                        if (coll2_Child.Entities.Count > 0)
                                                                                        {
                                                                                            foreach (Entity item3_Child in coll2_Child.Entities)
                                                                                            {
                                                                                                Entity obj_listItem_Child = service.Retrieve(item3_Child.LogicalName, item3_Child.Id, new ColumnSet(true)); ;

                                                                                                Guid productidorderchild = ((EntityReference)(obj_listItem_Child.Attributes["productid"])).Id;
                                                                                                Money amountchild = (Money)obj_listItem_Child.Attributes["amount"];
                                                                                                EntityReference uomchild = (EntityReference)obj_listItem_Child.Attributes["uomid"];

                                                                                                Entity new_Opp_Product_Child = new Entity("salesorderdetail");

                                                                                                new_Opp_Product_Child["productid"] = new EntityReference("product", productidorderchild);
                                                                                                new_Opp_Product_Child["priceperunit"] = amountchild;
                                                                                                new_Opp_Product_Child["salesorderid"] = new EntityReference(_order.LogicalName, _order.Id);
                                                                                                new_Opp_Product_Child["uomid"] = uomchild;
                                                                                                new_Opp_Product_Child["quantity"] = (decimal)1;
                                                                                                service.Create(new_Opp_Product_Child);
                                                                                                tracingService.Trace("Order Product created for Child-Child Product and to its Proce List");
                                                                                            }
                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            //No Price List Item Configured For Selected Price List  
                                                                                        }
                                                                                    }
                                                                                }
                                                                                else
                                                                                {
                                                                                    //Product Does not contain child products
                                                                                }
                                                                            }
                                                                        }
                                                                        #endregion
                                                                    }
                                                                }
                                                                // }
                                                                #endregion
                                                            }
                                                            else
                                                            {
                                                                #region For Upgrade
                                                                string Productname = _new_producttest.Attributes["name"].ToString();
                                                                OptionSetValue statuscode = (OptionSetValue)_new_producttest.Attributes["statuscode"];

                                                                if (statuscode.Value == 1)
                                                                {
                                                                    QueryExpression query2 = new QueryExpression("productpricelevel");
                                                                    query2.ColumnSet = new ColumnSet(true);
                                                                    FilterExpression filter = new FilterExpression(LogicalOperator.And);
                                                                    ConditionExpression cond2 = new ConditionExpression("pricelevelid", ConditionOperator.Equal, pricelist.Id);
                                                                    ConditionExpression cond3 = new ConditionExpression("productid", ConditionOperator.Equal, _new_producttest.Id);
                                                                    filter.AddCondition(cond2);
                                                                    filter.AddCondition(cond3);
                                                                    query2.Criteria.AddFilter(filter);
                                                                    EntityCollection coll2 = service.RetrieveMultiple(query2);
                                                                    tracingService.Trace("Price list item retrieval completed" + coll2.Entities.Count);
                                                                    if (coll2.Entities.Count > 0)
                                                                    {
                                                                        foreach (Entity obj_listItem in coll2.Entities)
                                                                        {
                                                                            tracingService.Trace("Create Order Product");
                                                                            //Entity obj_listItem = service.Retrieve(item3.LogicalName, item3.Id, new ColumnSet(true));
                                                                            Guid productidorder = ((EntityReference)(obj_listItem.Attributes["productid"])).Id;
                                                                            string productname = ((EntityReference)(obj_listItem.Attributes["productid"])).Name;
                                                                            Money amount = (Money)obj_listItem.Attributes["amount"];
                                                                            EntityReference uom = (EntityReference)obj_listItem.Attributes["uomid"];


                                                                            //Create Order Product Records
                                                                            if (productname.Contains("OTC"))
                                                                            {
                                                                                if (newOTC.Value > 0)
                                                                                {
                                                                                    tracingService.Trace("Product Contains OTC");
                                                                                    Entity new_Opp_Product = new Entity("salesorderdetail");
                                                                                    new_Opp_Product["priceperunit"] = newOTC;
                                                                                    new_Opp_Product["salesorderid"] = new EntityReference(_order.LogicalName, _order.Id);
                                                                                    new_Opp_Product["quantity"] = (decimal)1;
                                                                                    new_Opp_Product["isproductoverridden"] = true;
                                                                                    new_Opp_Product["baseamount"] = newOTC;
                                                                                    new_Opp_Product["productdescription"] = productname;
                                                                                    service.Create(new_Opp_Product);
                                                                                    tracingService.Trace("OTC Order Product got created");
                                                                                }
                                                                            }
                                                                            else
                                                                            {
                                                                                tracingService.Trace("Product does not Contains OTC");
                                                                                Entity new_Opp_Product = new Entity("salesorderdetail");
                                                                                new_Opp_Product["productid"] = new EntityReference("product", productidorder);
                                                                                new_Opp_Product["priceperunit"] = amount;
                                                                                new_Opp_Product["salesorderid"] = new EntityReference(_order.LogicalName, _order.Id);
                                                                                new_Opp_Product["uomid"] = uom;
                                                                                new_Opp_Product["quantity"] = (decimal)1;
                                                                                service.Create(new_Opp_Product);
                                                                                tracingService.Trace("RC Order Product got created");
                                                                            }
                                                                        }

                                                                        #region Check For Childs For Related Product
                                                                        if (_new_producttest.Attributes.Contains("alletech_chargetype"))
                                                                        {
                                                                            if (((OptionSetValue)(_new_producttest.Attributes["alletech_chargetype"])).Value == 569480000)
                                                                            {
                                                                                tracingService.Trace("If Charge Type is Package then retrieve child product related childs");
                                                                                QueryExpression query1_Child = new QueryExpression("alletech_productparent_productchild");
                                                                                query1_Child.ColumnSet = new ColumnSet(true);

                                                                                ConditionExpression cond1_Child = new ConditionExpression("productidone", ConditionOperator.Equal, _new_producttest.Id);

                                                                                query1_Child.Criteria.AddCondition(cond1_Child);
                                                                                EntityCollection coll1_Child = service.RetrieveMultiple(query1_Child);

                                                                                if (childproducts.Entities.Count > 0)
                                                                                {
                                                                                    foreach (Entity item1_Child in coll1_Child.Entities)
                                                                                    {
                                                                                        tracingService.Trace("Child product is having Child products");
                                                                                        Guid childchildpro = new Guid((item1.Attributes["productidtwo"].ToString()));
                                                                                        Entity _new_producttest_Child = service.Retrieve("product", childchildpro, new ColumnSet(true));


                                                                                        QueryExpression query2_Child = new QueryExpression("productpricelevel");
                                                                                        query2_Child.ColumnSet = new ColumnSet(true);
                                                                                        FilterExpression filter_Child = new FilterExpression(LogicalOperator.And);
                                                                                        ConditionExpression cond2_Child = new ConditionExpression("pricelevelid", ConditionOperator.Equal, pricelist.Id);
                                                                                        ConditionExpression cond3_Child = new ConditionExpression("productid", ConditionOperator.Equal, _new_producttest_Child.Id);
                                                                                        filter_Child.AddCondition(cond2_Child);
                                                                                        filter_Child.AddCondition(cond3_Child);
                                                                                        query2_Child.Criteria.AddFilter(filter_Child);
                                                                                        EntityCollection coll2_Child = service.RetrieveMultiple(query2_Child);
                                                                                        tracingService.Trace("Child-child product related procelist item retrieval completed");
                                                                                        if (coll2_Child.Entities.Count > 0)
                                                                                        {
                                                                                            foreach (Entity item3_Child in coll2_Child.Entities)
                                                                                            {
                                                                                                Entity obj_listItem_Child = service.Retrieve(item3_Child.LogicalName, item3_Child.Id, new ColumnSet(true)); ;

                                                                                                Guid productidorderchild = ((EntityReference)(obj_listItem_Child.Attributes["productid"])).Id;
                                                                                                string ChildProdName = ((EntityReference)(obj_listItem_Child.Attributes["productid"])).Name;
                                                                                                Money amountchild = (Money)obj_listItem_Child.Attributes["amount"];
                                                                                                EntityReference uomchild = (EntityReference)obj_listItem_Child.Attributes["uomid"];

                                                                                                Entity new_Opp_Product_Child = new Entity("salesorderdetail");
                                                                                                //Create Order Product Records
                                                                                                if (ChildProdName.Contains("OTC"))
                                                                                                {
                                                                                                    tracingService.Trace("Child Product Contains OTC");
                                                                                                    new_Opp_Product_Child["priceperunit"] = newOTC;
                                                                                                    new_Opp_Product_Child["salesorderid"] = new EntityReference(_order.LogicalName, _order.Id);
                                                                                                    new_Opp_Product_Child["quantity"] = (decimal)1;
                                                                                                    new_Opp_Product_Child["isproductoverridden"] = true;
                                                                                                    new_Opp_Product_Child["baseamount"] = newOTC;
                                                                                                    new_Opp_Product_Child["productdescription"] = ChildProdName;
                                                                                                    service.Create(new_Opp_Product_Child);
                                                                                                    tracingService.Trace("Child OTC Order Product got created");
                                                                                                }
                                                                                                else
                                                                                                {
                                                                                                    tracingService.Trace("Child Product doesnot Contains OTC");
                                                                                                    new_Opp_Product_Child["productid"] = new EntityReference("product", productidorderchild);
                                                                                                    new_Opp_Product_Child["priceperunit"] = amountchild;
                                                                                                    new_Opp_Product_Child["salesorderid"] = new EntityReference(_order.LogicalName, _order.Id);
                                                                                                    new_Opp_Product_Child["uomid"] = uomchild;
                                                                                                    new_Opp_Product_Child["quantity"] = (decimal)1;
                                                                                                    service.Create(new_Opp_Product_Child);
                                                                                                    tracingService.Trace("Child RC Order Product got created");
                                                                                                }
                                                                                            }
                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            //No Price List Item Configured For Selected Price List  
                                                                                        }
                                                                                    }
                                                                                }
                                                                                else
                                                                                {
                                                                                    //Product Does not contain child products
                                                                                }
                                                                            }
                                                                        }
                                                                        #endregion

                                                                    }
                                                                }

                                                                #endregion
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                //No Product selected for Corrosponding Price List  
                                            }
                                            if (Normal_useProduct_Id != Guid.Empty)
                                            {
                                                // _order["spectra_product"] = new EntityReference("product", Normal_useProduct_Id);
                                                // _order["spectra_getrelatedproducts"] = string.Empty;
                                                Entity order = new Entity("salesorder");
                                                order["spectra_productpkgcount"] = Convert.ToInt32(total_norml_prd_count);
                                                order.Id = _order.Id;
                                                service.Update(order);
                                            }
                                        }

                                    }
                                }
                                else
                                {
                                    throw new Exception("Account is not associated with Building");
                                }
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {

                throw new InvalidPluginExecutionException(ex.Message);
            }

        }

        public EntityCollection Getdetails(IOrganizationService service, ITracingService trace, string EntitySchema, string FieldSchema, Guid FieldValue)
        {
            trace.Trace("method Getdetails start Entity : " + EntitySchema);
            EntityCollection result = null;

            //Get Order Products Associated to Orders
            QueryExpression queryorderproduct = new QueryExpression(EntitySchema);
            queryorderproduct.ColumnSet = new ColumnSet(true);

            ConditionExpression cond = new ConditionExpression(FieldSchema, ConditionOperator.Equal, FieldValue);
            queryorderproduct.Criteria.AddCondition(cond);
            result = service.RetrieveMultiple(queryorderproduct);
            trace.Trace("method Getdetails End - Count of records : " + result.Entities.Count);
            return result;

        }

        public EntityCollection GetOTCdetails(IOrganizationService service, ITracingService trace, string EntitySchema, OptionSetValue type, Money baseotc, OptionSetValue tobandwidth, OptionSetValue frombandwidth)
        {
            trace.Trace("method Getdetails start Entity : " + EntitySchema + " " + type.Value + " " + baseotc.Value + " " + tobandwidth.Value + " " + frombandwidth.Value);
            EntityCollection result = null;

            //Get Order Products Associated to Orders
            QueryExpression queryorderproduct = new QueryExpression(EntitySchema);
            queryorderproduct.ColumnSet = new ColumnSet(true);

            ConditionExpression cond1 = new ConditionExpression("spectra_type", ConditionOperator.Equal, type.Value);
            ConditionExpression cond2 = new ConditionExpression("spectra_baseotc", ConditionOperator.Equal, baseotc.Value);
            ConditionExpression cond3 = new ConditionExpression("spectra_tobandwidthotc", ConditionOperator.Equal, tobandwidth.Value);
            ConditionExpression cond4 = new ConditionExpression("spectra_frombandwidthotc", ConditionOperator.Equal, frombandwidth.Value);
            queryorderproduct.Criteria.AddCondition(cond1);
            queryorderproduct.Criteria.AddCondition(cond2);
            queryorderproduct.Criteria.AddCondition(cond3);
            queryorderproduct.Criteria.AddCondition(cond4);
            result = service.RetrieveMultiple(queryorderproduct);
            trace.Trace("method Getdetails End - Count of records : " + result.Entities.Count);
            return result;

        }
    }
}
