using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace OrderDOA
{
    class OrderDOAHelper
    {

        public string getEmailBody(IOrganizationService service, string Approver, Entity order, Entity account, Entity sr, string custsegment, string billcyle, string arc, ITracingService trace)
        {

            string emailbody = null;
            try
            {

                trace.Trace("Email Body Method Started");
                string TempName = null; //used for all entity reference name values
                string TempName2 = null;//used for Building status and billing frequency.
                string OrderID = string.Empty;
                string Productname = string.Empty;
                string Reason = string.Empty;
                string SRM = string.Empty;
                string accountname = string.Empty;
                string CANID = string.Empty;
                string City = string.Empty;
                string activationdate = string.Empty;
                string PreviousProduct = string.Empty;
                string Businesssegment = string.Empty;
                string Typeofrequest = string.Empty;
                string Caseid = string.Empty;
                int uptype = 0;
                string opptysubtable = "";
                string SubType = string.Empty, SubSubType = string.Empty;

                trace.Trace("order details Started");
                //Order Details
                if (order != null)
                {
                    trace.Trace("order details are not null");
                    //Order ID
                    if (order.Attributes.Contains("ordernumber"))
                    {
                        OrderID = order.Attributes["ordernumber"].ToString();
                    }
                    //Product
                    if (order.Attributes.Contains("spectra_product"))
                    {
                        Productname = ((EntityReference)order.Attributes["spectra_product"]).Name;
                    }
                    //Reason
                    if (order.Attributes.Contains("spectra_discountreason"))
                    {
                        Reason = order.Attributes["spectra_discountreason"].ToString();
                    }

                    uptype = order.GetAttributeValue<OptionSetValue>("prioritycode").Value;

                    trace.Trace("order details completed");
                }
                //Account Details
                trace.Trace("account details are started");
                if (account != null)
                {
                    trace.Trace("account details are not null");
                    //SRM
                    if (account.Attributes.Contains("spectra_servicerelationshipmanagerid"))
                    {
                        SRM = ((EntityReference)account.Attributes["spectra_servicerelationshipmanagerid"]).Name;
                    }
                    //account name
                    if (account.Attributes.Contains("name"))
                    {
                        accountname = account.Attributes["name"].ToString();
                    }
                    //CAN ID
                    if (account.Attributes.Contains("alletech_accountid"))
                    {
                        CANID = account.Attributes["alletech_accountid"].ToString();
                    }
                    //City
                    if (account.Attributes.Contains("alletech_city"))
                    {
                        City = ((EntityReference)account.Attributes["alletech_city"]).Name;
                    }
                    //activationdate
                    if (account.Attributes.Contains("alletech_activationdate"))
                    {
                        activationdate = account.Attributes["alletech_activationdate"].ToString();
                    }
                    //PreviousProduct
                    if (account.Attributes.Contains("alletech_product"))
                    {
                        PreviousProduct = ((EntityReference)account.Attributes["alletech_product"]).Name;
                    }
                    trace.Trace("account details are completed");

                }

                //Case Details
                if (sr != null)
                {
                    trace.Trace("sr details started");
                    //Businesssegment
                    if (sr.Attributes.Contains("alletech_businesssegment"))
                    {
                        Businesssegment = ((EntityReference)sr.Attributes["alletech_businesssegment"]).Name;
                    }
                    //Typeofrequest Upgrade or Down grade
                    if (sr.Attributes.Contains("alletech_subdisposition"))
                    {
                        Typeofrequest = ((EntityReference)sr.Attributes["alletech_subdisposition"]).Name;
                    }
                    //Typeofrequest Upgrade or Down grade
                    if (sr.Attributes.Contains("alletech_subdisposition"))
                    {
                        Typeofrequest = ((EntityReference)sr.Attributes["alletech_subdisposition"]).Name;
                        SubType = ((EntityReference)sr.Attributes["alletech_subdisposition"]).Name.ToString();
                    }
                    if (sr.Attributes.Contains("alletech_disposition"))
                    {
                        SubSubType = ((EntityReference)sr.Attributes["alletech_disposition"]).Name.ToString();
                    }
                    //Case ID
                    if (sr.Attributes.Contains("alletech_caseidcrm"))
                    {
                        Caseid = sr.Attributes["alletech_caseidcrm"].ToString();
                    }
                    trace.Trace("SR details End");
                }

                #region Constructing HTML
                emailbody = "<html><body><div>";
                emailbody += "Hi " + Approver + ",<br/><br/>";
                //type
                if (uptype == 111260000)
                    emailbody += "Approval of "+ Typeofrequest + " for order Id " + OrderID + ",<br/>";
                else if (uptype == 111260001)
                    emailbody += "Approval of Temporary Upgradation for order " + OrderID + ",<br/>";
                else if (uptype == 111260002)
                    emailbody += "Approval of B2C Downgrade for order " + OrderID + ",<br/>";

                emailbody += "<p style = 'color: black; font-family: ' Arial',sans-serif; font-size: 10pt;'><b> Order Details </b></p>";
                emailbody += "<table class='MsoNormalTable' style='width:547pt;font-family: Arial,sans-serif; font-size: 9pt; border-collapse: collapse;' border='0' cellspacing='0' cellpadding='0'><tbody>";

                #region Row 1
                //if (Oppty.Attributes.Contains("spectra_customersegmentcode"))
                //    TempName = GetCustomersegment(Oppty.GetAttributeValue<OptionSetValue>("spectra_customersegmentcode").Value);
                //else
                //    TempName = null;


                if (custsegment == null || custsegment == string.Empty)
                {
                    emailbody += @"<tr style='height: 15pt;'>
                    <td width='239' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 179pt; height: 15pt; ' colspan='2'>
                        <p align='center' text-align: center;'><b>Business Segment</b></p>
                    </td>
                    <td width='131' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 98pt; height: 15pt; '>
                        <p align='center' text-align: center;'>" + Businesssegment + @"</p>
                    </td>
                    <td width='119' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 89pt; height: 15pt; '>
                        <p><b>Account Manager</b></p>
                    </td>
                    <td width='241' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 181pt; height: 15pt; ' colspan='2'>
                        <p align='center' text-align: center;'>" + SRM + @"</p>
                    </td>                   
                </tr>";

                }
                else
                {
                    emailbody += @"<tr style ='height: 15pt;'>
                                <td width ='142' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 106.35pt; height: 15pt; '>
                                <p><b> Business Segment </b></p></td>
                                <td width ='97' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 72.65pt; height: 15pt; '>

                                <p align ='center' text-align: 'center;'>" + Businesssegment + @"</p></td>

                                <td width ='131' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 98pt; height: 15pt; '>
                                <p><b> Customer Segment </b></p></td>
                                <td width ='119' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 89pt; height: 15pt; '>
                                
                                <p align ='center' text-align: 'center;'>" + custsegment + @"</p></td>
                                
                                <td width ='93' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 70pt; height: 15pt; '>
                                <p><b> SRM Name </b></p></td>
                                <td width ='148' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 111pt; height: 15pt; '>
                                
                                <p align ='center' text-align: 'center;'>" + SRM + @"</p></td></tr>";
                }

                #endregion

                #region New Row added
                emailbody += @"<tr style='height: 15pt;'>
                    <td width='239' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 179pt; height: 15pt; ' colspan='2'>
                        <p align='center' text-align: center;'><b>Type</b></p>
                    </td>
                    <td width='131' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 98pt; height: 15pt; '>
                        <p align='center' text-align: center;'>" + SubType + @"</p>
                    </td>
                    <td width='119' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 89pt; height: 15pt; '>
                        <p><b>Sub Type</b></p>
                    </td>
                    <td width='241' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 181pt; height: 15pt; ' colspan='2'>
                        <p align='center' text-align: center;'>" + SubSubType + @"</p>
                    </td>                 
                </tr>";
                #endregion

                #region Row 2
                //if (Oppty.Attributes.Contains("alletech_companynamebusiness"))
                //    TempName = Oppty.GetAttributeValue<string>("alletech_companynamebusiness");
                //else
                //    TempName = null;

                emailbody += @"<tr style ='height: 15pt;'>
                                <td width ='142' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 106.35pt; height: 15pt; '>
                                <p><b> Customer Name </b></p></td>
                                <td width ='97' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 72.65pt; height: 15pt; '>

                                <p align ='center' text-align: 'center;'>" + accountname + @"</p></td>

                                <td width ='131' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 98pt; height: 15pt; '>
                                <p><b> Service ID </b></p></td>
                                <td width ='119' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 89pt; height: 15pt; '>
                                
                                <p align ='center' text-align: 'center;'>" + CANID + @"</p></td>
                                
                                <td width ='93' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 70pt; height: 15pt; '>
                                <p><b> City </b></p></td>
                                <td width ='148' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 111pt; height: 15pt; '>
                                
                                <p align ='center' text-align: 'center;'>" + City + @"</p></td></tr>";
                #endregion

                #region Row 3
                //if (Oppty.Attributes.Contains("alletech_area"))
                //    TempName = Oppty.GetAttributeValue<EntityReference>("alletech_area").Name;
                //else
                //    TempName = null;

                emailbody += @"<tr style ='height: 15pt;'>
                                <td width ='142' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 106.35pt; height: 15pt; '>
                                <p><b> Customer Activation date </b></p></td>
                                <td width ='97' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 72.65pt; height: 15pt; '>

                                <p align ='center' text-align: 'center;'>" + activationdate + @"</p></td>

                                <td width ='131' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 98pt; height: 15pt; '>
                                <p><b> Current Plan </b></p></td>
                                <td width ='119' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 89pt; height: 15pt; '>
                                
                                <p align ='center' text-align: 'center;'>" + PreviousProduct + @"</p></td>";


                if (uptype == 111260000)//Permanent
                {
                    emailbody += @"<td width ='93' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 70pt; height: 15pt; '>
                                <p><b> Current ARC </b></p></td>
                                <td width ='148' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 111pt; height: 15pt; '>
                                
                                <p align ='center' text-align: 'center;'>" + arc + @"</p></td></tr>";
                }
                else if (uptype == 111260001) //Temp
                {
                    DateTime strtdate = order.GetAttributeValue<DateTime>("spectra_effectivedate");
                    DateTime enddate = order.GetAttributeValue<DateTime>("submitdate");

                    double days = (enddate - strtdate).TotalDays;

                    emailbody += @"<td width ='93' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 70pt; height: 15pt; '>
                                <p><b> Number of days </b></p></td>
                                <td width ='148' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 111pt; height: 15pt; '>
                                
                                <p align ='center' text-align: 'center;'>" + days + @"</p></td></tr>";
                }
                else if(uptype == 111260002)//Home 
                {
                    emailbody += @"<td width ='93' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 70pt; height: 15pt; '>
                                <p><b> Current MRC </b></p></td>
                                <td width ='148' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 111pt; height: 15pt; '>
                                
                                <p align ='center' text-align: 'center;'>" + arc + @"</p></td></tr>";
                }
                #endregion

                if (uptype == 111260000)//Permanent
                {
                    #region Checking OpptyProds
                    string[] Opptyvalues = new string[4];
                    TempName = null; TempName2 = null;

                    string[] ordprodattr = { "productid", "priceperunit", "manualdiscountamount", "extendedamount", "productdescription" };
                    EntityCollection ordProds = GetResultsByAttribute(service, "salesorderdetail", "salesorderid", order.Id.ToString(), ordprodattr);

                    if (ordProds.Entities.Count > 0)
                    {
                        foreach (Entity ordprod in ordProds.Entities)
                        {
                            string prodname = null;
                            Entity prod = null;
                            string[] prodattr = { "alletech_billingcycle", "alletech_plantype", "alletech_chargetype" };
                            if (ordprod.Attributes.Contains("productid"))
                            {
                                prodname = ordprod.GetAttributeValue<EntityReference>("productid").Name;
                                prod = GetResultByAttribute(service, "product", "productid", ordprod.GetAttributeValue<EntityReference>("productid").Id.ToString(), prodattr);
                            }
                            else
                            {
                                prodname = ordprod.GetAttributeValue<string>("productdescription");
                                prod = GetResultByAttribute(service, "product", "name", prodname, prodattr);
                            }

                            try
                            {
                                int plan = prod.GetAttributeValue<OptionSetValue>("alletech_plantype").Value;
                                int charge = prod.GetAttributeValue<OptionSetValue>("alletech_chargetype").Value;

                                decimal price = 0, extnamt = 0, percentAge = 0;
                                //if (prodname.EndsWith("RC") || prodname.EndsWith("OTC") || plan == 569480002) //addon plan type
                                if (charge == 569480001)// || charge == 569480002 || plan == 569480002) //addon plan type
                                {
                                    price = ordprod.GetAttributeValue<Money>("priceperunit").Value;
                                    Opptyvalues[0] = price.ToString();

                                    if (ordprod.Contains("manualdiscountamount"))
                                        Opptyvalues[1] = ordprod.GetAttributeValue<Money>("manualdiscountamount").Value.ToString();
                                    else
                                        Opptyvalues[1] = "0";

                                    extnamt = ordprod.GetAttributeValue<Money>("extendedamount").Value;
                                    Opptyvalues[2] = extnamt.ToString();

                                    if (extnamt < price)
                                    {
                                        percentAge = (price - extnamt) / price * 100;
                                        percentAge = decimal.Round(percentAge, 2, MidpointRounding.AwayFromZero);
                                    }

                                    Opptyvalues[3] = percentAge.ToString() + "%";

                                    opptysubtable += "<tr style='height: 15pt;'>";
                                    opptysubtable += "<td width='142' style='border-width: 0px 1pt 1pt; border-style: none solid solid; padding: 0in 5.4pt; width: 106.35pt;'> ";
                                    opptysubtable += "<p align='center' text-align: center;'><b>" + prodname + "</b></p>";
                                    opptysubtable += "</td>";
                                    opptysubtable += "<td width='97' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none;  padding: 0in 5.4pt; width: 72.65pt;'> ";
                                    opptysubtable += "<p align='center' text-align: center;'>" + Opptyvalues[0] + @"</p>";
                                    opptysubtable += "</td>";
                                    opptysubtable += "<td width='131' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none;  padding: 0in 5.4pt; width: 98pt;'> ";
                                    opptysubtable += "<p align='center' text-align: center;'>" + Opptyvalues[1] + @"</p>";
                                    opptysubtable += "</td>";
                                    opptysubtable += "<td width='119' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none;  padding: 0in 5.4pt; width: 89pt;'> ";
                                    opptysubtable += "<p align='center' text-align: center;'>" + Opptyvalues[2] + @"</p>";
                                    opptysubtable += "</td>";
                                    opptysubtable += "<td width='241' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none;  padding: 0in 5.4pt; width: 181pt;' colspan='2'> ";
                                    opptysubtable += "<p align='center' text-align: center;'>" + Opptyvalues[3] + @"</p>";
                                    opptysubtable += "</td>";
                                    opptysubtable += "</tr>";
                                }
                                else
                                {
                                    TempName = prodname;

                                    if (prod.Attributes.Contains("alletech_billingcycle"))
                                        TempName2 = prod.GetAttributeValue<EntityReference>("alletech_billingcycle").Name;
                                }
                            }
                            catch (Exception ex)
                            {
                                opptysubtable += "Exception In building sub table : " + ex.Message;
                            }
                        }
                    }
                    #endregion
                }

                #region Row 4 & 5
                emailbody += @"<tr style='height: 15pt;'>
                    <td width='239' style='border-width: 0px 1pt 1pt; border-style: none solid solid; border-color: rgb(0, 0, 0) black windowtext windowtext; padding: 0in 5.4pt; width: 179pt; height: 15pt; ' colspan='2'>
                        <p align='center' text-align: center;'><b>Requested (new) Plan</b></p>
                    </td>
                    <td width='131' style='border-width: 0px 0px 1pt; border-style: none none solid; border-color: rgb(0, 0, 0) rgb(0, 0, 0) windowtext; padding: 0in 5.4pt; width: 98pt; height: 15pt; '>
                        <p align='center' text-align: center;'>" + Productname + @"</p>
                    </td>
                    <td width='119' style='border-width: 0px 1pt 1pt; border-style: none solid solid; padding: 0in 5.4pt; width: 89pt; height: 15pt; '>
                        <p><b>Billing Frequency</b></p>
                    </td>
                    <td width='241' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none; border-color: rgb(0, 0, 0) black windowtext rgb(0, 0, 0); padding: 0in 5.4pt; width: 181pt; height: 15pt; ' colspan='2'>
                        <p align='center' text-align: center;'>" + billcyle + @"</p>
                    </td>
                </tr>";

                emailbody += @"<tr style='height: 15pt;'>
                    <td width='239' style='border-width: 0px 1pt 1pt; border-style: none solid solid; border-color: rgb(0, 0, 0) black windowtext windowtext; padding: 0in 5.4pt; width: 179pt; height: 15pt; ' colspan='2'>
                        <p align='center' text-align: center;'><b>Type of Request</b></p>
                    </td>
                    <td width='131' style='border-width: 0px 0px 1pt; border-style: none none solid; border-color: rgb(0, 0, 0) rgb(0, 0, 0) windowtext; padding: 0in 5.4pt; width: 98pt; height: 15pt; '>
                        <p align='center' text-align: center;'>" + Typeofrequest + @"</p>
                    </td>
                    <td width='119' style='border-width: 0px 1pt 1pt; border-style: none solid solid; padding: 0in 5.4pt; width: 89pt; height: 15pt; '>
                        <p><b>SR No.</b></p>
                    </td>
                    <td width='241' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none; border-color: rgb(0, 0, 0) black windowtext rgb(0, 0, 0); padding: 0in 5.4pt; width: 181pt; height: 15pt; ' colspan='2'>
                        <p align='center' text-align: center;'>" + Caseid + @"</p>
                    </td>
                </tr>";
                #endregion

                if (uptype == 111260000)//Permanent
                {
                    #region Row 6
                    emailbody += @"<tr style='height: 15pt;'>
                    <td width='142' style='border-width: 0px 1pt 1pt; border-style: none solid solid; padding: 0in 5.4pt; width: 106.35pt; height: 15pt; '>
                        <p align='center' text-align: center;'><b>Product Component</b></p>
                    </td>
                    <td width='97' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none;  padding: 0in 5.4pt; width: 72.65pt; height: 15pt; '>
                        <p align='center' text-align: center;'><b>Floor Price</b></p>
                    </td>
                    <td width='131' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none;  padding: 0in 5.4pt; width: 98pt; height: 15pt; '>
                        <p align='center' text-align: center;'><b>Discount</b></p>
                    </td>
                    <td width='119' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none;  padding: 0in 5.4pt; width: 89pt; height: 15pt; '>
                        <p align='center' text-align: center;'><b>Discounted Price</b></p>
                    </td>
                    <td width='241' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none;  padding: 0in 5.4pt; width: 181pt; height: 15pt; ' colspan='2'>
                        <p align='center' text-align: center;'><b>Price Deviation %</b></p>
                    </td>
                </tr>";
                    #endregion

                    emailbody += opptysubtable;

                    #region Row 9

                    emailbody += @"<tr style='height: 33pt;'>
                    <td width='142' style='border-width: 0px 1pt 1pt; border-style: none solid solid; padding: 0in 5.4pt; width: 106.35pt; height: 33pt; '>
                        <p><b>Remarks</b></p>
                    </td>
                    <td width='588' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none;  padding: 0in 5.4pt; width: 440.65pt; height: 33pt; ' colspan='5'>
                        <p align='center' text-align: center;'>" + Reason + @"</p>
                    </td>
                </tr>";
                    #endregion
                }

                emailbody += "</tbody></table>";
                emailbody += "<br /> Note: To approve from email, please reply to this email with just 'Approve'/'Approved'.";
                emailbody += "<br />       To reject from email, please reply to this email with just 'Reject'/'Rejected'.";
                emailbody += "<br /><br /> Regards,<br /> Spectra Team <br />";
                emailbody += "</div></body></html>";

                #endregion
                //}
            }
            catch (Exception ex)
            {
                emailbody = "Exception In creating  body : " + ex.Message;
            }
            return emailbody;
        }

        public EntityCollection getApprovalConfig(IOrganizationService service, string appConfigNameType, string prefix)//, decimal percentAge)
        {
            QueryExpression query = new QueryExpression("spectra_approvalconfig");
            query.ColumnSet = new ColumnSet("spectra_approver", "spectra_name", "spectra_orderby", "spectra_minpercentage", "spectra_maxpercentage", "spectra_quantity");//spectra_percentage
            query.Criteria.AddCondition(new ConditionExpression("spectra_name", ConditionOperator.Equal, prefix + appConfigNameType.ToUpper()));
            query.Criteria.AddCondition(new ConditionExpression("spectra_productsegment", ConditionOperator.Null));
            query.Orders.Add(new OrderExpression("spectra_minpercentage", OrderType.Ascending));
            query.Orders.Add(new OrderExpression("spectra_orderby", OrderType.Ascending));
            return service.RetrieveMultiple(query);
        }

        public EntityCollection getApprovalConfigB2B(IOrganizationService service, string appConfigNameType, string prefix)//, decimal percentAge)
        {
            QueryExpression query = new QueryExpression("spectra_approvalconfig");
            query.ColumnSet = new ColumnSet("spectra_approver", "spectra_name", "spectra_orderby", "spectra_minpercentage", "spectra_maxpercentage", "spectra_quantity");//spectra_percentage
            query.Criteria.AddCondition(new ConditionExpression("spectra_name", ConditionOperator.Equal, prefix + appConfigNameType.ToUpper()));
            query.Criteria.AddCondition(new ConditionExpression("spectra_productsegment", ConditionOperator.NotNull));
            query.Orders.Add(new OrderExpression("spectra_minpercentage", OrderType.Ascending));
            query.Orders.Add(new OrderExpression("spectra_orderby", OrderType.Ascending));
            return service.RetrieveMultiple(query);
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

                if (entityName != "systemuser")
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

        public Entity GetResultByAttribute(IOrganizationService _service, string entityName, string attrName, string attrValue, string[] column)
        {
            try
            {
                Entity result = null;
                QueryExpression query = new QueryExpression(entityName);

                query.NoLock = true;
                query.ColumnSet.AddColumns(column);
                query.Criteria.AddCondition(attrName, ConditionOperator.Equal, attrValue);

                //query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

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

        public EntityCollection GetResultsByAttribute(IOrganizationService _service, string entityName, string attrName, string attrValue, string[] column)
        {
            try
            {
                QueryExpression query = new QueryExpression(entityName);
                query.NoLock = true;
                query.ColumnSet.AddColumns(column);
                query.Criteria.AddCondition(attrName, ConditionOperator.Equal, attrValue);
                //query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

                EntityCollection resultcollection = _service.RetrieveMultiple(query);

                return resultcollection;

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

    }
}
