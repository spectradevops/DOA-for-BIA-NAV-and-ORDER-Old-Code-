using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrderDOA
{
    public class UpdateApprovedPercentageInOppProd : CodeActivity
    {
        [Input("Opportunity")]
        [RequiredArgument]
        [ReferenceTarget("salesorder")]
        public InArgument<EntityReference> Opportunity { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService traceService = executionContext.GetExtension<ITracingService>();
            //Obtain WorkflwoContext from the executionContext.
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            //Obtain the organization service reference.
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            if (Opportunity.Get(executionContext).Id != null)
            {
                EntityCollection entCollOppProd = getOppProducts(service, Opportunity.Get(executionContext).Id);

                foreach (Entity entOppProd in entCollOppProd.Entities)
                {
                    if (entOppProd.Contains("productid"))
                    {
                        EntityReference prodId = (EntityReference)entOppProd["productid"];
                        if (prodId.Name.ToLower().Contains("_ipaddress_") || prodId.Name.ToLower().EndsWith("rc") || prodId.Name.ToLower().EndsWith("otc"))
                        {
                            Entity entProd = service.Retrieve(prodId.LogicalName, prodId.Id, new ColumnSet("alletech_businesssegmentlookup", "alletech_grossplaninvoicevalueinr"));
                            if (entProd.Contains("alletech_businesssegmentlookup") && ((EntityReference)entProd["alletech_businesssegmentlookup"]).Name.ToLower() == "business")
                            {
                                decimal percentAge = 0;
                                decimal extendedAmt = ((Money)entOppProd["extendedamount"]).Value;

                                if (prodId.Name.ToLower().Contains("_ipaddress_"))
                                {
                                    ////string searchData = "_IPADDRESS_";

                                    //var cnt = prodId.Name.Substring(prodId.Name.IndexOf(searchData) + searchData.Length);
                                    //cnt = Regex.Replace(cnt, "[^0-9]+", string.Empty);
                                    //Int64 count = Convert.ToInt64(cnt);

                                    //percentAge = extendedAmt / count;

                                    //EntityCollection entCollAppConfig = getApprovalConfig(service, "IPADDRESS");//, percentAge);
                                    //foreach (Entity entAppConfig in entCollAppConfig.Entities)
                                    //{
                                    //    if ((entAppConfig.Contains("spectra_quantity") && count >= Convert.ToInt64(entAppConfig["spectra_quantity"])) || !entAppConfig.Contains("spectra_quantity"))
                                    //    {
                                    //        if (entAppConfig.Contains("spectra_minpercentage") && entAppConfig.Contains("spectra_maxpercentage")
                                    //            && (percentAge >= ((decimal)entAppConfig["spectra_minpercentage"])) && (percentAge <= ((decimal)entAppConfig["spectra_maxpercentage"])))
                                    //        {
                                    //            if (entAppConfig.Contains("spectra_orderby") && entAppConfig.Contains("spectra_approver"))
                                    //            {
                                    //                KeyValuePair<int, Guid> apprDuplicate = approvals.FirstOrDefault(a => a.Key == Convert.ToInt16(entAppConfig.FormattedValues["spectra_orderby"]));
                                    //                if (apprDuplicate.Equals(new KeyValuePair<int, Guid>()))
                                    //                    approvals.Add(Convert.ToInt16(entAppConfig.FormattedValues["spectra_orderby"]), ((EntityReference)entAppConfig["spectra_approver"]).Id);
                                    //            }
                                    //        }
                                    //    }
                                    //}
                                }

                                else if (prodId.Name.ToLower().EndsWith("rc") || prodId.Name.ToLower().EndsWith("otc"))
                                {
                                    if (entProd.Contains("alletech_grossplaninvoicevalueinr"))
                                    {
                                        decimal floorDisc = ((Money)entProd["alletech_grossplaninvoicevalueinr"]).Value;
                                        if (extendedAmt < floorDisc)
                                        {
                                            percentAge = (floorDisc - extendedAmt) / floorDisc * 100;
                                            percentAge = decimal.Round(percentAge, 2, MidpointRounding.AwayFromZero);

                                            //if (entOppProd.Contains("spectra_approvalrequried") && (Boolean)entOppProd["spectra_approvalrequried"])
                                            if (!entOppProd.Contains("spectra_approvedpercentage") || (entOppProd.Contains("spectra_approvedpercentage") && (decimal)entOppProd["spectra_approvedpercentage"] < percentAge))
                                            {
                                                Entity entOppProdUpdate = new Entity(entOppProd.LogicalName);
                                                entOppProdUpdate.Id = entOppProd.Id;
                                                entOppProdUpdate["spectra_approvedpercentage"] = percentAge;
                                                entOppProdUpdate["spectra_approvalrequried"] = false;
                                                service.Update(entOppProdUpdate);
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

        public EntityCollection getOppProducts(IOrganizationService service, Guid oppId)
        {
            QueryExpression query = new QueryExpression("salesorderdetail");
            query.ColumnSet = new ColumnSet("extendedamount", "productid", "spectra_approvalrequried", "spectra_approvedpercentage");
            query.Criteria.AddCondition(new ConditionExpression("salesorderid", ConditionOperator.Equal, oppId));

            return service.RetrieveMultiple(query);
        }

    }
}
