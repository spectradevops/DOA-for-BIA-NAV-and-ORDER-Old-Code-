using Microsoft.Crm.Sdk.Messages;
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
    public class DeactivateAllApprovalsOnReject : CodeActivity
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

            if (context.PrimaryEntityName.ToLower() == "spectra_approval")
            {
                if (Opportunity.Get(executionContext) != null)
                {
                    EntityReference oppId = Opportunity.Get(executionContext);

                    QueryExpression query = new QueryExpression("spectra_approval");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria.AddCondition(new ConditionExpression("spectra_orderid", ConditionOperator.Equal, oppId.Id));
                    query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                    EntityCollection entCollApproval = service.RetrieveMultiple(query);
                    //throw new Exception("Count of Approvals "+entCollApproval.Entities.Count);
                    foreach (Entity entApproval in entCollApproval.Entities)
                    {
                        //entApproval["statecode"] = new OptionSetValue(1);
                        //entApproval["statuscode"] = new OptionSetValue(2);
                        //service.Update(entApproval);
                        SetStateRequest staReq = new SetStateRequest()
                        {
                            EntityMoniker = new EntityReference("spectra_approval", entApproval.Id),
                            State = new OptionSetValue(1),
                            Status = new OptionSetValue(2)
                        };
                        service.Execute(staReq);
                    }
                }
            }
        }
    }
}
