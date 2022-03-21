using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Feasibility_DOA
{
    public class FeasibilityAproveTlinkProductUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                //if (context.Depth == 1)
                //{
                Entity feasib = (Entity)context.InputParameters["Target"];
                if (feasib.Attributes.Contains("spectra_approvalstatus"))
                {
                    if (feasib.Attributes["spectra_approvalstatus"] != null)
                    {
                        tracingService.Trace("spectra_approvalstatus");
                        if (feasib.GetAttributeValue<OptionSetValue>("spectra_approvalstatus").Value == 1)
                        {
                            #region Feasibility
                            Entity _feasibility = service.Retrieve("alletech_feasibility", feasib.Id, new ColumnSet("alletech_routetype", "alletech_product", "alletech_opportunity"));
                            if (_feasibility != null)
                            {
                                tracingService.Trace("alletech_feasibility");
                                if (_feasibility.Attributes.Contains("alletech_opportunity"))
                                {
                                    tracingService.Trace("alletech_opportunity");
                                    string fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                  <entity name='opportunityproduct'>
                                                    <attribute name='productid' />
                                                    <attribute name='extendedamount' />
                                                    <attribute name='opportunityproductid' />
                                                    <order attribute='productid' descending='false' />
                                                    <filter type='and'>
                                                      <condition attribute='opportunityid' operator='eq' value='" + ((EntityReference)_feasibility.Attributes["alletech_opportunity"]).Id + @"' />
                                                    </filter>
                                                    <link-entity name='product' from='productid' to='productid' visible='false' link-type='outer' alias='oppprod'>
                                                      <attribute name='alletech_plantype' />
                                                      <attribute name='alletech_chargetype' />                                                        
                                                    </link-entity>
                                                  </entity>
                                                </fetch>";

                                    EntityCollection oppProdCol = service.RetrieveMultiple(new FetchExpression(fetch));
                                    if (oppProdCol.Entities.Count > 0)
                                    {
                                        tracingService.Trace("oppProdCol.Entities.Count");
                                        foreach (Entity oppprd in oppProdCol.Entities)
                                        {
                                            if (oppprd.Attributes.Contains("productid"))
                                            {
                                                int plantype = ((OptionSetValue)oppprd.GetAttributeValue<AliasedValue>("oppprod.alletech_plantype").Value).Value;
                                                int chargetype = ((OptionSetValue)oppprd.GetAttributeValue<AliasedValue>("oppprod.alletech_chargetype").Value).Value;
                                                if (plantype == 569480001 && chargetype == 569480000)
                                                {
                                                    service.Delete("opportunityproduct", oppprd.Id);
                                                }
                                            }
                                        }
                                        if (_feasibility.Attributes.Contains("alletech_product"))
                                        {
                                            string productName = ((EntityReference)_feasibility.Attributes["alletech_product"]).Name.ToString() + "_T";
                                            string productFetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                                  <entity name='product'>
                                                                    <attribute name='alletech_plantype' />
                                                                    <attribute name='defaultuomid' />
                                                                    <attribute name='productid' />
                                                                    <order attribute='productnumber' descending='false' />
                                                                    <filter type='and'>
                                                                      <condition attribute='name' operator='eq' value='" + productName + @"' />
                                                                      <condition attribute='alletech_businesssegmentlookup' operator='eq' uiname='Business' uitype='alletech_businesssegment' value='{B6D61BE4-ACCB-E411-942D-842B2BA0F44F}' />                                                                      
                                                                    </filter>
                                                                    <link-entity name='alletech_subsegment' from='alletech_subsegmentid' to='alletech_subbusinesssegment' alias='ag'>
                                                                          <filter type='and'>
                                                                            <condition attribute='alletech_name' operator='ne' value='SDWAN' />
                                                                          </filter>
                                                                    </link-entity>
                                                                    <link-entity name='alletech_productsegment' from='alletech_productsegmentid' to='alletech_productsegment' alias='ah'>
                                                                          <filter type='and'>
                                                                            <condition attribute='alletech_name' operator='eq' value='Secured Managed Internet' />
                                                                          </filter>
                                                                    </link-entity>
                                                                  </entity>
                                                                </fetch>";
                                            EntityCollection prodColle = service.RetrieveMultiple(new FetchExpression(productFetch));
                                            if (prodColle.Entities.Count > 0)
                                            {
                                                Guid unit_ID = prodColle.Entities[0].GetAttributeValue<EntityReference>("defaultuomid").Id;
                                                Entity unitDetail = service.Retrieve("uom", unit_ID, new ColumnSet("name"));
                                                if (unitDetail != null)
                                                {
                                                    if (unitDetail.Attributes.Contains("name"))
                                                    {
                                                        var unitName = unitDetail.Contains("name") ? unitDetail.GetAttributeValue<string>("name") : "";
                                                        var planType = prodColle.Entities[0].Contains("alletech_plantype") ? prodColle.Entities[0].GetAttributeValue<OptionSetValue>("alletech_plantype").Value.ToString() : "0";

                                                        if (planType == "569480001")//Normal
                                                        {
                                                            Entity opportunityProduct = new Entity("opportunityproduct");
                                                            opportunityProduct["productid"] = new EntityReference("product", prodColle.Entities[0].Id);
                                                            opportunityProduct["spectra_existingproduct"] = new EntityReference("product", prodColle.Entities[0].Id);
                                                            opportunityProduct["opportunityid"] = new EntityReference("opportunity", ((EntityReference)_feasibility.Attributes["alletech_opportunity"]).Id);
                                                            opportunityProduct["isproductoverridden"] = false;
                                                            opportunityProduct["quantity"] = 1m;
                                                            opportunityProduct["manualdiscountamount"] = new Money(0);
                                                            if (!string.IsNullOrEmpty(unitName))
                                                            {
                                                                opportunityProduct["uomid"] = new EntityReference("uom", unit_ID);
                                                            }
                                                            Guid opD = service.Create(opportunityProduct);

                                                            if (opD != null && opD != Guid.Empty)
                                                            {
                                                                Entity opportunity = new Entity("opportunity");
                                                                opportunity.Id = ((EntityReference)_feasibility.Attributes["alletech_opportunity"]).Id;
                                                                //opportunity["alletech_product"] = new EntityReference("product", prodColle.Entities[0].Id);
                                                                opportunity["alletech_getrelatedproducts"] = true;
                                                                opportunity["alletech_redundancyrequired"] = false;
                                                                opportunity["spectra_lastmiletype"] = new OptionSetValue(2);
                                                                service.Update(opportunity);

                                                                string feasibFetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                                                  <entity name='alletech_feasibility'>
                                                                                    <attribute name='alletech_feasibilityid' />
                                                                                    <attribute name='alletech_feasibilityidd' />
                                                                                    <attribute name='createdon' />
                                                                                    <order attribute='alletech_feasibilityidd' descending='false' />
                                                                                    <filter type='and'>
                                                                                      <condition attribute='alletech_opportunity' operator='eq' value='" + ((EntityReference)_feasibility.Attributes["alletech_opportunity"]).Id + @"' />
                                                                                    </filter>
                                                                                  </entity>
                                                                                </fetch>";
                                                                EntityCollection fetchColl = service.RetrieveMultiple(new FetchExpression(feasibFetch));
                                                                if (fetchColl.Entities.Count > 0)
                                                                {
                                                                    foreach (Entity FSB in fetchColl.Entities)
                                                                    {
                                                                        Entity feasibUpdate = new Entity("alletech_feasibility");
                                                                        feasibUpdate.Id = FSB.Id;
                                                                        if (FSB.GetAttribute‌​‌​Value<bool>("alletech_routetype") == false)
                                                                        {                                                                          
                                                                            feasibUpdate["alletech_redundent"] = false;
                                                                            feasibUpdate["alletech_product"] = new EntityReference("product", prodColle.Entities[0].Id);
                                                                            service.Update(feasibUpdate);
                                                                        }
                                                                        if (FSB.GetAttribute‌​‌​Value<bool>("alletech_routetype") == true)
                                                                        { 
                                                                            feasibUpdate["alletech_redundent"] = false;
                                                                            service.Update(feasibUpdate);
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
                        }
                    }
                    #endregion
                }
                // }
            }
            catch (Exception ex)
            {

                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
