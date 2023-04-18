using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Presales_DOA
{
    public class CreateApproval : IPlugin
    {
        private string configData = string.Empty;
        private Dictionary<string, string> globalConfig = new Dictionary<string, string>();
        public CreateApproval(string unsecureConfig)
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
            string entTo = string.Empty, entcc1 = string.Empty, entcc2 = string.Empty, entcc3 = string.Empty;
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            try
            {
                entTo = this.GetValueForKey("To");
                entcc1 = this.GetValueForKey("CC1");
                entcc2 = this.GetValueForKey("CC2");
                entcc3 = this.GetValueForKey("CC3");
                tracingService.Trace("Service Created");
                Entity primaryEntity = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet(true));
                if (primaryEntity.Contains("onl_opportunityid"))
                {
                    Entity Oppo = service.Retrieve("opportunity", ((EntityReference)primaryEntity.Attributes["onl_opportunityid"]).Id, new ColumnSet("alletech_oppurtunityid", "ownerid"));
                    string Oppoid = Oppo.Attributes["alletech_oppurtunityid"].ToString();
                    string OwnerId = Oppo.GetAttributeValue<EntityReference>("ownerid").Id.ToString();
                    int presalesstatus = primaryEntity.Contains("onl_status") ? primaryEntity.GetAttributeValue<OptionSetValue>("onl_status").Value : 0; // Pending for approval
                    if (presalesstatus == 122050003)
                    {
                        Entity To = getUserentityreference(service, entTo);
                        Entity CC1 = getUserentityreference(service, entcc1);
                        Entity CC2 = getUserentityreference(service, entcc2);
                        Entity CC3 = getUserentityreference(service, entcc3);

                        #region Create approval record
                        Entity entApproval = new Entity("spectra_approval");
                        entApproval["spectra_name"] = "Pre-Sales approval";
                        if (To != null)
                            entApproval["ownerid"] = new EntityReference("systemuser", To.Id);
                        entApproval["spectra_approvalrequesteddate"] = DateTime.Now;
                        entApproval["statecode"] = new OptionSetValue(0);
                        entApproval["statuscode"] = new OptionSetValue(111260000);
                        entApproval["spectra_preslaes"] = new EntityReference("onl_presalestask", context.PrimaryEntityId);
                        Guid approvalId = service.Create(entApproval);
                        #endregion

                        #region Email Create
                        Entity user = service.Retrieve("systemuser", To.Id, new ColumnSet("fullname"));
                        String UserName = user.Attributes["fullname"].ToString();
                        Entity entEmail = new Entity("email");
                        entEmail["subject"] = "Pending for your approval #" + approvalId.ToString().ToUpper() + "#";
                        string emailbody = getEmailBody("Pending for your approval #" + approvalId.ToString().ToUpper() + "#", Oppoid);
                        entEmail["description"] = "Hi " + UserName + ",\n" + emailbody;

                        Entity ToParty = new Entity("activityparty");
                        ToParty["partyid"] = new EntityReference("systemuser", user.Id);
                        Entity[] entToList = { ToParty };
                        entEmail["to"] = entToList;

                        Entity CCParty1 = null;
                        Entity CCParty2 = null;
                        Entity CCParty3 = null;
                        Entity CCParty4 = null;
                        if (CC1 != null)
                        {
                            CCParty1 = new Entity("activityparty");
                            CCParty1["partyid"] = new EntityReference("systemuser", CC1.Id);
                        }
                        if (CC2 != null)
                        {
                            CCParty2 = new Entity("activityparty");
                            CCParty2["partyid"] = new EntityReference("systemuser", CC2.Id);
                        }
                        if (CC2 != null)
                        {
                            CCParty2 = new Entity("activityparty");
                            CCParty2["partyid"] = new EntityReference("systemuser", CC2.Id);
                        }
                        if (CC3 != null)
                        {
                            CCParty3 = new Entity("activityparty");
                            CCParty3["partyid"] = new EntityReference("systemuser", CC3.Id);
                        }
                        CCParty4 = new Entity("activityparty");
                        CCParty4["partyid"] = new EntityReference("systemuser", Guid.Parse(OwnerId)); // Opportunity Owner

                        Entity[] entccList = { CCParty1, CCParty2, CCParty3, CCParty4 };
                        entEmail["cc"] = entccList;

                        Entity Queue = GetResultByAttribute(service, "queue", "name", "DOA Approval", "queueid");
                        if (Queue != null)
                        {
                            Entity entFrom = new Entity("activityparty");
                            entFrom["partyid"] = new EntityReference("queue", Queue.Id);
                            Entity[] entFromList = { entFrom };
                            entEmail["from"] = entFromList;
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException("DOA approval not available");
                        }
                        entEmail["regardingobjectid"] = new EntityReference("spectra_approval", approvalId);
                        Guid _emailId = service.Create(entEmail);
                        #endregion

                        #region Attachment
                        if (_emailId != Guid.Empty)
                        {
                            QueryExpression query_docs = new QueryExpression("annotation");
                            query_docs.ColumnSet = new ColumnSet(true);
                            ConditionExpression cond_docs = new ConditionExpression("objectid", ConditionOperator.Equal, primaryEntity.Id.ToString());
                            query_docs.Criteria.AddCondition(cond_docs);
                            EntityCollection annotationCollection = service.RetrieveMultiple(query_docs);
                            if (annotationCollection != null && annotationCollection.Entities.Count > 0)
                            {
                                foreach (Entity annotationItem in annotationCollection.Entities)
                                {

                                    Entity attachment = new Entity("activitymimeattachment");
                                    attachment["subject"] = "Policy";
                                    string fileName = annotationItem.Attributes["filename"].ToString();
                                    attachment["filename"] = fileName;
                                    byte[] fileStream = Encoding.ASCII.GetBytes(annotationItem.Attributes["documentbody"].ToString());
                                    //attachments[0].Attributes[AnnotationEntities.DocumentBody]
                                    attachment["body"] = Convert.ToBase64String(Convert.FromBase64String((string)annotationItem["documentbody"]));
                                    attachment["mimetype"] = annotationItem.Attributes["mimetype"].ToString();
                                    attachment["attachmentnumber"] = 1;
                                    attachment["objectid"] = new EntityReference("email", _emailId);
                                    attachment["objecttypecode"] = "email";
                                    Guid guid = service.Create(attachment);
                                }
                            }
                            SendEmailRequest sendEmailRequest = new SendEmailRequest
                            {
                                EmailId = _emailId,
                                TrackingToken = "",
                                IssueSend = true
                            };
                            SendEmailResponse sendEmailresp = (SendEmailResponse)service.Execute(sendEmailRequest);
                        }
                        #endregion
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
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

        public Entity getUserentityreference(IOrganizationService service, string useremailID)
        {
            Entity result = null;

            try
            {
                string userFetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='systemuser'>
                                    <attribute name='fullname' />
                                    <attribute name='businessunitid' />
                                    <attribute name='title' />
                                    <attribute name='address1_telephone1' />
                                    <attribute name='positionid' />
                                    <attribute name='systemuserid' />
                                    <order attribute='fullname' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='internalemailaddress' operator='eq' value='" + useremailID + @"' />
                                      <condition attribute='isdisabled' operator='eq' value='0' />
                                    </filter>
                                  </entity>
                                </fetch>";
                EntityCollection userCollection = service.RetrieveMultiple(new FetchExpression(userFetch));
                if (userCollection.Entities.Count > 0)
                {
                    return userCollection.Entities[0];
                }
                else
                {
                    return result;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }

        }
        public string getEmailBody(string Subject, string Oppoid)
        {
            string content = string.Empty;
            content += "<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8' /><meta http-equiv='X-UA-Compatible' content='IE=edge' /><meta name='viewport' content='width=device-width, initial-scale=1.0' />";
            content += "</head>";
            content += "<body>";
            content += "<br /><br />Approval of solution document for opportunity ID :" + Oppoid + ".<br /><br />";
            content += "Note: To approve from email , Please reply to this email with just 'Approve/Approved'<br /><br />";
            content += "To reject from email, Please reply to this email with  just 'Reject/Rejected'.<br /><br />";
            content += "Team Spectra<br />";
            content += "#MakeLifeBetter";
            content += "</body></html>";
            return content;
        }

        public Entity GetResultByAttribute(IOrganizationService _service, string entityName, string attrName, string attrValue, string column)
        {
            try
            {
                Entity result = null;
                QueryExpression query = new QueryExpression(entityName);
                query.NoLock = true;

                if (column == "all")
                    query.ColumnSet.AllColumns = true;
                else
                    query.ColumnSet.AddColumns(column);

                query.Criteria.AddCondition(attrName, ConditionOperator.Equal, attrValue);

                if (entityName != "opportunity" && entityName != "systemuser")
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

                EntityCollection resultcollection = _service.RetrieveMultiple(query);

                if (resultcollection.Entities.Count > 0)
                    return resultcollection.Entities[0];
                else
                    return result;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
