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
                Entity feasib = (Entity)context.InputParameters["Target"];
                if (feasib.Attributes.Contains("spectra_approvalstatus"))
                {
                    if (feasib.GetAttributeValue<OptionSetValue>("alletech_installationstatusnew").Value == 1)
                    {
                        Entity _feasibility = service.Retrieve("alletech_feasibility", feasib.Id, new ColumnSet("alletech_routetype", "alletech_product", "alletech_opportunity"));
                        if (_feasibility != null)
                        {
                            if (_feasibility.Attributes.Contains("alletech_opportunity"))
                            {
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
                                                                    <attribute name='name' />
                                                                    <attribute name='defaultuomid' />
                                                                    <attribute name='productid' />
                                                                    <order attribute='productnumber' descending='false' />
                                                                    <filter type='and'>
                                                                      <condition attribute='name' operator='eq' value='" + productName + @"' />
                                                                      <condition attribute='alletech_businesssegmentlookup' operator='eq' uiname='Business' uitype='alletech_businesssegment' value='{B6D61BE4-ACCB-E411-942D-842B2BA0F44F}' />
                                                                      <condition attribute='alletech_subbusinesssegment' operator='ne' uiname='SDWAN' uitype='alletech_subsegment' value='{857B1258-5F33-EA11-80EE-000D3AF224B9}' />
                                                                      <condition attribute='alletech_productsegment' operator='eq' uiname='Secured Managed Internet' uitype='alletech_productsegment' value='{A91D2A5C-D654-EC11-8127-000D3AC9A7B3}' />
                                                                    </filter>
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
                                                                    feasibUpdate["alletech_redundent"] = false;
                                                                    feasibUpdate["alletech_product"] = new EntityReference("product", prodColle.Entities[0].Id);
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
            catch (Exception ex)
            {

                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
