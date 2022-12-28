using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace SDWAN_PreSales_DOA
{
    public class PreSalesDOA : CodeActivity
    {

        [Input("Type")]
        [RequiredArgument]
        public InArgument<string> Type { get; set; }

        [Output("Guid")]
        public OutArgument<string> ApprovalGUID { get; set; }

        [Output("OpportunityID")]
        public OutArgument<string> OpportunityID { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            string type = Type.Get(executionContext);
            try
            {
                if (type == "approval")
                {
                    ApprovalGUID.Set(executionContext, context.PrimaryEntityId.ToString());

                    Entity approval = service.Retrieve("spectra_approval", context.PrimaryEntityId, new ColumnSet("spectra_presalestask"));
                    if (approval.Attributes.Contains("spectra_presalestask"))
                    {
                        string presalesID = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='onl_presalestask'>
                                            <attribute name='onl_presalestaskid' />
                                            <attribute name='onl_name' />
                                            <attribute name='createdon' />
                                            <order attribute='onl_name' descending='false' />
                                            <filter type='and'>
                                              <condition attribute='onl_presalestaskid' operator='eq' value='" + ((EntityReference)approval.Attributes["spectra_presalestask"]).Id + @"' />
                                            </filter>
                                            <link-entity name='opportunity' from='opportunityid' to='onl_opportunityid' visible='false' link-type='outer' alias='OPP'>
                                              <attribute name='alletech_oppurtunityid' />
                                            </link-entity>
                                          </entity>
                                        </fetch>";
                        EntityCollection presalesCollection = service.RetrieveMultiple(new FetchExpression(presalesID));
                        if(presalesCollection.Entities.Count > 0)
                        {

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
