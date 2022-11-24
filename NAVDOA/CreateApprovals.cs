using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Web.Script.Serialization;
using System.Xml;

namespace NAVDOA
{
    public class CreateApprovals : IPlugin
    {
        IPluginExecutionContext context;
        IOrganizationService service;
        ITracingService tracingService;

        private string configData = string.Empty;
        private Dictionary<string, string> globalConfig = new Dictionary<string, string>();

        public CreateApprovals(string unsecureConfig)
        {
            if (string.IsNullOrEmpty(unsecureConfig))
            {
                throw new InvalidPluginExecutionException("Unsecure configuration missing.");
            }

            this.configData = unsecureConfig;
            this.ReadUnSecuredConfig(this.configData);
        }
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
                #region Email Parameters 24 Nov 2022
                string URL = string.Empty, Authkey = string.Empty, Action = string.Empty, from1 = string.Empty, to = string.Empty, cc = string.Empty, subject = string.Empty, content = string.Empty;
               URL = this.GetValueForKey("ML_URL");
                Authkey = this.GetValueForKey("ML_AuthKey");
                Action = this.GetValueForKey("ML_Action");
                from1 = this.GetValueForKey("ML_From");
                #endregion

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

                
                string emailbody = helper.getEmailBody(service,postImg, CHO.Name,Type);

                #region Creating EMAIL Activity
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

                #region New Logic on 24 Nov 2022
                subject = "Pending for your approval #" + approvalId.ToString().ToUpper() + "#";
                content = emailbody.ToString();

                Entity entApprover1 = service.Retrieve("systemuser", CHO.Id, new ColumnSet("internalemailaddress"));

                if (entApprover1.Attributes.Contains("internalemailaddress"))
                    to = entApprover1.Attributes["internalemailaddress"].ToString();

                #region ML API Calling
                string integrationCheck = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                              <entity name='alletech_integrationlog'>
                                                                <attribute name='alletech_name' />
                                                                <attribute name='createdon' />
                                                                <attribute name='alletech_regardingentity' />
                                                                <attribute name='alletech_integrationwith' />
                                                                <attribute name='alletech_comment' />
                                                                <attribute name='alletech_integrationlogid' />
                                                                <order attribute='createdon' descending='true' />
                                                                <filter type='and'>
                                                                  <condition attribute='alletech_entityguid' operator='eq' value='" + approvalId + @"' />
                                                                </filter>
                                                              </entity>
                                                            </fetch>";
                EntityCollection integrColl = service.RetrieveMultiple(new FetchExpression(integrationCheck));
                if (integrColl.Entities.Count == 0)
                {
                    var result = string.Empty;
                    string json = string.Empty;
                    var httpWebRequest = (HttpWebRequest)WebRequest.Create(URL);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = "POST";
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                    using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                    {
                        // #region CRM parameters mapping with ML parameters
                        SendingEmail app = new SendingEmail();
                        app.Authkey = Authkey;
                        app.Action = Action;
                        app.from = from1;
                        app.to = to;
                        app.cc = cc;
                        app.subject = subject;
                        app.content = content;
                        json = new JavaScriptSerializer().Serialize(app);
                        streamWriter.Write(json);
                        streamWriter.Flush();
                        streamWriter.Close();
                    }
                    //tracingService.Trace("Json generated");
                    var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        result = streamReader.ReadToEnd();
                        if (result.Contains("success"))
                        {
                            Entity log = new Entity("alletech_integrationlog");
                            log["alletech_name"] = "NAV Approval Email Pushed to ML: " + approvalId.ToString();
                            log["alletech_integrationwith"] = new OptionSetValue(4);
                            log["alletech_regardingentity"] = "alletech_feasibility";
                            log["alletech_entityguid"] = approvalId.ToString();
                            log["alletech_request"] = json.ToString();
                            log["alletech_responce"] = result.ToString();
                            service.Create(log);
                        }
                        else
                        {
                            Entity log = new Entity("alletech_integrationlog");
                            log["alletech_name"] = "NAV Approval Email Pushed to ML: " + approvalId.ToString();
                            log["alletech_integrationwith"] = new OptionSetValue(4);
                            log["alletech_regardingentity"] = "alletech_feasibility";
                            log["alletech_entityguid"] = approvalId.ToString();
                            log["alletech_request"] = json.ToString();
                            log["alletech_responce"] = result.ToString();
                            service.Create(log);
                        }
                    }
                }
                #endregion
                #endregion
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error in creating approval or email : " + ex.Message);
            }
        }
        private string GetValueForKey(string keyName)
        {
            string valueString = string.Empty;
            try
            {
                if (this.globalConfig.ContainsKey(keyName))
                {
                    valueString = this.globalConfig[keyName];
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return valueString;
        }
        private void ReadUnSecuredConfig(string localConfig)
        {
            string key = string.Empty;
            try
            {
                this.globalConfig = new Dictionary<string, string>();
                if (!string.IsNullOrEmpty(localConfig))
                {
                    XmlDocument doc = new XmlDocument();

                    doc.LoadXml(localConfig);

                    foreach (XmlElement entityNode in doc.SelectNodes("/appSettings/add"))
                    {
                        key = entityNode.GetAttribute("key").ToString();
                        this.globalConfig.Add(entityNode.GetAttribute("key").ToString(), entityNode.GetAttribute("value").ToString());
                    }
                }
            }

            catch (Exception e)
            {
                throw e;
            }
        }
    }
    public class SendingEmail
    {
        public string Authkey { get; set; }
        public string Action { get; set; }
        public string from { get; set; }
        public string to { get; set; }
        public string cc { get; set; }
        public string subject { get; set; }
        public string content { get; set; }

    }
}
