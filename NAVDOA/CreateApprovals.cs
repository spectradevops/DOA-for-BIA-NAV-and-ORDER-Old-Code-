using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System;

namespace NAVDOA
{
    public class CreateApprovals : IPlugin
    {
        IPluginExecutionContext context;
        IOrganizationService service;
        ITracingService tracingService;
        public void Execute(IServiceProvider serviceProvider)
        {
            //Context = Info passed to the plugin at runtime
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));

            //Service = access to data for modification
            service = factory.CreateOrganizationService(context.UserId);
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                NAVDOAHelper helper = new NAVDOAHelper();
                Entity target = (Entity)context.InputParameters["Target"];

                if (context.MessageName == "Update")
                {
                    Entity postImg = context.PostEntityImages["PostImage"];

                    tracingService.Trace("In update");

                    if (target.Attributes.Contains("alletech_atthenearestmanhole") && target.GetAttributeValue<bool>("alletech_atthenearestmanhole"))
                    {
                        if (postImg.Attributes.Contains("spectra_zone"))
                        {
                            tracingService.Trace("contains ZONE");
                            Entity zone = helper.GetResultByAttribute(service, "spectra_zone", "spectra_zoneid", postImg.GetAttributeValue<EntityReference>("spectra_zone").Id.ToString(), "spectra_cho");

                            if (zone.Attributes.Contains("spectra_cho"))
                            {
                                EntityReference CHO = zone.GetAttributeValue<EntityReference>("spectra_cho");

                                tracingService.Trace("CHO : " + CHO.Name);

                                CreateApprovalandEMail(CHO, postImg,helper,"WCR");

                                if (context.PrimaryEntityName == "alletech_wcr")
                                {
                                    #region updating WCR consumption status
                                    Entity wcr = new Entity("alletech_wcr");
                                    wcr.Id = postImg.Id;
                                    wcr["statuscode"] = new OptionSetValue(111260004);//waiting for approval
                                    service.Update(wcr);
                                    #endregion
                                }
                            }
                            else
                                throw new InvalidPluginExecutionException("CHO not mapped in Zone");
                        }
                        else
                            throw new InvalidPluginExecutionException("WCR not mapped with Zone");
                    }

                    else if(target.Attributes.Contains("spectra_speedtestresultsshowntocustomer") && target.GetAttributeValue<bool>("spectra_speedtestresultsshowntocustomer"))
                    {
                        if (postImg.Attributes.Contains("spectra_zone"))
                        {
                            tracingService.Trace("contains ZONE");
                            Entity zone = helper.GetResultByAttribute(service, "spectra_zone", "spectra_zoneid", postImg.GetAttributeValue<EntityReference>("spectra_zone").Id.ToString(), "spectra_cho");

                            if (zone.Attributes.Contains("spectra_cho"))
                            {
                                EntityReference CHO = zone.GetAttributeValue<EntityReference>("spectra_cho");

                                tracingService.Trace("CHO : " + CHO.Name);

                                CreateApprovalandEMail(CHO, postImg,helper,"IR");

                                if (context.PrimaryEntityName == "alletech_installationform")
                                {
                                    #region updating IR consumption status
                                    Entity IR = new Entity("alletech_installationform");
                                    IR.Id = postImg.Id;
                                    IR["statuscode"] = new OptionSetValue(111260004);//waiting for approval
                                    service.Update(IR);
                                    #endregion
                                }
                            }
                            else
                                throw new InvalidPluginExecutionException("CHO not mapped in Zone");
                        }
                        else
                            throw new InvalidPluginExecutionException("IR not mapped with Zone");
                    }
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("NAV DOA plugin : " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        public void CreateApprovalandEMail(EntityReference CHO,Entity postImg,NAVDOAHelper helper,string Type)
        {
            try
            {
                #region Create approval record
                Entity entApproval = new Entity("spectra_approval");
                entApproval["spectra_name"] = "CHO approval";
                entApproval["ownerid"] = new EntityReference("systemuser", CHO.Id);
                entApproval["spectra_approvalrequesteddate"] = DateTime.Now;
                entApproval["statecode"] = new OptionSetValue(0);
                entApproval["statuscode"] = new OptionSetValue(111260000);

                tracingService.Trace("context.PrimaryEntityName : " + context.PrimaryEntityName);

                if (context.PrimaryEntityName == "alletech_wcr")
                    entApproval["spectra_wcr"] = new EntityReference("alletech_wcr", postImg.Id);
                else
                    entApproval["spectra_installationreport"] = new EntityReference("alletech_installationform", postImg.Id);

                tracingService.Trace("approval before create");
                Guid approvalId = service.Create(entApproval);

                tracingService.Trace("approval created");
                #endregion

                #region Creating EMAIL Activity
                string emailbody = helper.getEmailBody(service,postImg, CHO.Name,Type);

                Entity entEmail = new Entity("email");
                entEmail["subject"] = "Pending for your approval #" + approvalId.ToString().ToUpper() + "#";
                entEmail["description"] = emailbody;
                
                Entity entTo = new Entity("activityparty");
                entTo["partyid"] = new EntityReference("systemuser", CHO.Id);
                Entity[] entToList = { entTo };
                entEmail["to"] = entToList;

                Entity Queue = helper.GetResultByAttribute(service, "queue", "name", "DOA Approval", "queueid");

                if (Queue != null)
                {
                    Entity entFrom = new Entity("activityparty");
                    entFrom["partyid"] = new EntityReference("queue", Queue.Id);
                    Entity[] entFromList = { entFrom };
                    entEmail["from"] = entFromList;
                }
                else
                    throw new InvalidPluginExecutionException("DOA approval not available");

                entEmail["regardingobjectid"] = new EntityReference("spectra_approval", approvalId);
                Guid emailId = service.Create(entEmail);
                tracingService.Trace("Email created");

                //Send email
                SendEmailRequest sendEmailReq = new SendEmailRequest()
                {
                    EmailId = emailId,
                    IssueSend = true
                };
                SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);
                #endregion
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error in creating approval or email : " + ex.Message);
            }
        }

    }
}
