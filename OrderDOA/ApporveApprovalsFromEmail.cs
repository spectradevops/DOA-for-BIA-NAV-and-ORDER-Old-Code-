using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OrderDOA
{
    public class ApporveApprovalsFromEmail : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = (IOrganizationService)factory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService trace = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            if (context.PrimaryEntityName.ToLower() == "email")
            {
                if (context.MessageName.ToLower() == "create")
                {
                    trace.Trace("Create");

                    if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                    {
                        Entity entTraget = (Entity)context.InputParameters["Target"];
                        //Entity entTraget = service.Retrieve("email", context.PrimaryEntityId, new ColumnSet(true));
                        trace.Trace("Target");

                        if (entTraget.Contains("directioncode") && !((Boolean)entTraget["directioncode"]))
                        {
                            trace.Trace("incoming");
                            if (entTraget.Contains("regardingobjectid"))
                            {
                                trace.Trace("object checking");
                                EntityReference objID = (EntityReference)entTraget["regardingobjectid"];
                                if (objID.LogicalName.ToLower() == "spectra_approval")
                                {
                                    trace.Trace("Ok");
                                    Entity entApproval = service.Retrieve(objID.LogicalName, objID.Id, new ColumnSet("ownerid","spectra_orderid"));
                                    if (entTraget.Contains("from") && entApproval.Attributes.Contains("spectra_orderid"))
                                    {
                                        trace.Trace("from");
                                        EntityCollection entFromList = (EntityCollection)entTraget["from"];
                                        foreach (Entity entForm in entFromList.Entities)
                                        {
                                            trace.Trace(((EntityReference)entForm["partyid"]).Id.ToString());
                                            trace.Trace("Owner");
                                            trace.Trace(((EntityReference)entApproval["ownerid"]).Id.ToString());

                                            if (entForm.Contains("partyid") && ((EntityReference)entForm["partyid"]).Id.ToString() == ((EntityReference)entApproval["ownerid"]).Id.ToString())
                                            {
                                                trace.Trace("checking");

                                                string emailBody = entTraget["description"].ToString();
                                                
                                                string body = Regex.Replace(emailBody, "<.*?>", String.Empty).ToLower();
                                                
                                                string[] emailcontent = body.Split(new string[] { "from:" }, StringSplitOptions.None);

                                                if (emailcontent[0].Contains("approved") || emailcontent[0].Contains("approve"))
                                                {
                                                    trace.Trace("updating");
                                                    
                                                        entApproval["spectra_approveddate"] = DateTime.Now;
                                                        entApproval["statuscode"] = new OptionSetValue(111260001);
                                                        service.Update(entApproval);
                                                    
                                                    break;
                                                }
                                                else if (emailcontent[0].Contains("reject") || emailcontent[0].Contains("rejected"))
                                                {
                                                    entApproval["statuscode"] = new OptionSetValue(111260002);
                                                    entApproval["spectra_rejecteddate"] = DateTime.Now;
                                                    service.Update(entApproval);
                                                    break;
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
