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
        public string getEmailBody(IOrganizationService service, string Approver, Entity order, Entity account, Entity sr, string custsegment, string billcyle, string arc, ITracingService trace, string subject, string approvalID)
        {

            string content = null;
            try
            {

                string TempName = null; //used for all entity reference name values
                string TempName2 = null;//used for Building status and billing frequency.
                //string content = string.Empty;
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
                //Order Details
                if (order != null)
                {
                    //trace.Trace("order details are not null");
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

                    //trace.Trace("order details completed");
                }
                //Account Details
                //trace.Trace("account details are started");
                if (account != null)
                {
                    //trace.Trace("account details are not null");
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
                    //trace.Trace("account details are completed");

                }

                //Case Details
                if (sr != null)
                {
                    //trace.Trace("sr details started");
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
                    //trace.Trace("SR details End");

                }
                #region HTML Head
                content += "<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8' /><meta http-equiv='X-UA-Compatible' content='IE=edge' /><meta name='viewport' content='width=device-width, initial-scale=1.0' />";
                #region CSS
                content += @"<style>:root {
  --color-grey: #cccccc;
}
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}
html {
  font-size: 16px;
}

body {
  font-family: 'Roboto', sans-serif;
  background-color: #f5f5f5;
}
p {
  font-size: 1rem;
  line-height: 1.5;
  margin: unset;
}

div.container {
  width: 100%;
  max-width: 648px;
  margin: 0 auto;
  padding: 0 15px;
}
.fw-semibold {
  font-weight: 500;
}
.fw-bold {
  font-weight: 700;
}
.text-center {
  text-align: center;
}
.text-left {
  text-align: left;
}
.text-right {
  text-align: right;
}
.uppercase {
  text-transform: uppercase;
}
.bg-black {
  background-color: black !important;
  color: white;
}
table tr:nth-child(even) {
  background-color: #cccccc !important;
}
table.slip {
  width: 100%;
  border-collapse: collapse;
  border: 1px solid #666666;
}
table.slip tr td {
  padding: 0.25rem 0.75rem;
  border: 1px solid #666666;
}

@media (max-width: 768px) {
  html {
    font-size: 14px;
  }
  table.slip tr td {
    padding: 0.25rem;
  }
  p {
    font-size: 0.9rem;
  }
  div.container {
    padding: 0 10px;
  }
}</style>
";
                #endregion
                content += "</head>";
                #endregion
                #region HTML Body
                content += "<body><div class='container'><table class='slip'><!-- first row main header --><tr class='bg-black'><td colspan='5'>";
                if (uptype == 111260000)
                    content += "<p>" + subject +"</p> </td></tr>";
                else if (uptype == 111260001)
                    content += "<p>" + subject + "</p> </td></tr>";
                else if (uptype == 111260002)
                    content += "<p>Approval of B2C Downgrade for order " + approvalID + "</p> </td></tr>";
                else if (uptype == 111260003)
                    content += "<p>" + subject + "</p> </td></tr>";
                content += "<!-- Details --><!-- Order ID -->";
                content += "<tr style='background: #cccccc!important;'><td colspan='2'><p class='text-left'>Order Id:</p></td><td colspan='3'><p class='text-left'>" + OrderID + "</p></td></tr>";
                content += "<!-- Segment -->";
                if (custsegment == null || custsegment == string.Empty)
                {
                    content += "<tr><td colspan='2'><p class='text-left'>Segment</p></td><td colspan='3'><p class='text-left'>" + Businesssegment + "</p></td></tr>";
                    content += "<!-- Account Manager -->";
                    content += @"<tr style='background: #cccccc!important;'>
          <td colspan='2'>
            <p class='text-left'>Account Manager</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + SRM + "</p></td></tr>";
                }
                else
                {
                    content += "<tr><td colspan='2'><p class='text-left'>Segment</p></td><td colspan='3'><p class='text-left'>" + Businesssegment + "/" + custsegment + "</p></td></tr>";
                    content += @"<tr style='background: #cccccc!important;'>
          <td colspan='2'>
            <p class='text-left'>Account Manager</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + SRM + "</p></td></tr>";
                }
                content += @"<tr>
          <td colspan='2'>
            <p class='text-left'>Type</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + SubType + "</p></td></tr>";
                content += @"<tr style='background: #cccccc!important;'>
          <td colspan='2'>
            <p class='text-left'>Sub Type</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + SubSubType + "</p></td></tr>";
                content += @"<tr>
          <td colspan='2'>
            <p class='text-left'>Customer Name</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + accountname + "</p></td></tr>";
                content += @"<tr style='background: #cccccc!important;'>
          <td colspan='2'>
            <p class='text-left'>Service ID</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + CANID + "</p></td></tr>";
                content += @"<tr>
          <td colspan='2'>
            <p class='text-left'>City</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + City + "</p></td></tr>";
                content += @"<tr style='background: #cccccc!important;'>
          <td colspan='2'>
            <p class='text-left'>Customer Activation date</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + activationdate + "</p></td></tr>";
                content += @"<tr>
          <td colspan='2'>
            <p class='text-left'>Current Plan</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + PreviousProduct + "</p></td></tr>";

                if (uptype == 111260000 || uptype == 111260003)//Permanent || BIA to MBIA added on 17-03-2023
                {
                    content += @"<tr style='background: #cccccc!important;'>
          <td colspan='2'>
            <p class='text-left'>Current ARC</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + arc + "</p></td></tr>";
                }
                else if (uptype == 111260001) //Temp
                {
                    DateTime strtdate = order.GetAttributeValue<DateTime>("spectra_effectivedate");
                    DateTime enddate = order.GetAttributeValue<DateTime>("submitdate");

                    double days = (enddate - strtdate).TotalDays;
                    content += @"<tr>
          <td colspan='2'>
            <p class='text-left'>Number of days</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + days + "</p></td></tr>";
                }
                else if (uptype == 111260002)//Home 
                {
                    content += @"<tr style='background: #cccccc!important;'>
          <td colspan='2'>
            <p class='text-left'>Current MRC</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + arc + "</p></td></tr>";
                }
                if (uptype == 111260000 || uptype == 111260003)//Permanent || BIA to MBIA added on 17-03-2023
                {
                    #region Checking OpptyProds
                    string[] Opptyvalues = new string[4];
                    TempName = null; TempName2 = null;

                    string[] ordprodattr = { "productid", "priceperunit", "manualdiscountamount", "extendedamount", "productdescription" };
                    EntityCollection ordProds = GetResultsByAttribute(service, "salesorderdetail", "salesorderid", order.Id.ToString(), ordprodattr);

                    if (ordProds.Entities.Count > 0)
                    {
                        int count = ordProds.Entities.Count;
                        for (int i = 0; i < count; i++)
                        {
                            string prodname = null;
                            Entity prod = null;
                            string[] prodattr = { "alletech_billingcycle", "alletech_plantype", "alletech_chargetype" };
                            if (ordProds.Entities[i].Attributes.Contains("productid"))
                            {
                                prodname = ordProds.Entities[i].GetAttributeValue<EntityReference>("productid").Name;
                                prod = GetResultByAttribute(service, "product", "productid", ordProds.Entities[i].GetAttributeValue<EntityReference>("productid").Id.ToString(), prodattr);
                            }
                            else
                            {
                                prodname = ordProds.Entities[i].GetAttributeValue<string>("productdescription");
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
                                    price = ordProds.Entities[i].GetAttributeValue<Money>("priceperunit").Value;
                                    Opptyvalues[0] = price.ToString("0.#");

                                    if (ordProds.Entities[i].Contains("manualdiscountamount"))
                                        Opptyvalues[1] = ordProds.Entities[i].GetAttributeValue<Money>("manualdiscountamount").Value.ToString("0.#");
                                    else
                                        Opptyvalues[1] = "0";

                                    extnamt = ordProds.Entities[i].GetAttributeValue<Money>("extendedamount").Value;
                                    Opptyvalues[2] = extnamt.ToString("0.#");

                                    if (extnamt < price)
                                    {
                                        percentAge = (price - extnamt) / price * 100;
                                        percentAge = decimal.Round(percentAge, 2, MidpointRounding.AwayFromZero);
                                    }

                                    Opptyvalues[3] = percentAge.ToString() + "%";
                                    if (i % 2 == 0)
                                    {
                                        opptysubtable += @"<tr style='background: #cccccc!important;'><td><p class='text-left'>" + prodname + "</p></td><td><p class='text-right'>" + Opptyvalues[0] + "</p></td><td><p class='text-right'>" + Opptyvalues[1] + "</p></td><td><p class='text-right'>" + Opptyvalues[2] + "</p></td><td><p class='text-right'>" + Opptyvalues[3] + "</p></td></tr>";
                                    }
                                    else
                                    {
                                        opptysubtable += @"<tr ><td><p class='text-left'>" + prodname + "</p></td><td><p class='text-right'>" + Opptyvalues[0] + "</p></td><td><p class='text-right'>" + Opptyvalues[1] + "</p></td><td><p class='text-right'>" + Opptyvalues[2] + "</p></td><td><p class='text-right'>" + Opptyvalues[3] + "</p></td></tr>";

                                    }
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
                content += @"<tr>
          <td colspan='2'>
            <p class='text-left'>Requested (new) Plan</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + Productname + "</p></td></tr>";
                content += @"<tr style='background: #cccccc!important;'>
          <td colspan='2'>
            <p class='text-left'>Billing Frequency</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + billcyle + "</p></td></tr>";
                content += @"<tr>
          <td colspan='2'>
            <p class='text-left'>Type of Request</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + Typeofrequest + "</p></td></tr>";
                content += @"<tr style='background: #cccccc!important;'>
          <td colspan='2'>
            <p class='text-left'>SR No.</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + Caseid + "</p></td></tr>";
                if (uptype == 111260000 || uptype == 111260003)//Permanent || BIA to MBIA added on 17-03-2023
                {
                    content += @"<!-- empty row  -->
        <tr>
          <td colspan='5'>&nbsp;</td>
        </tr>";
                    content += @"<!-- Approval header -->
        <tr class='bg-grey' style='background: #cccccc!important;'>
          <td colspan='5'>
            <p class='text-left fw-bold'>Approval:</p>
          </td>
        </tr><!-- Approval components header -->
        <tr>
          <td>
            <p class='text-center'><b>Component</b></p>
          </td>
          <td>
            <p class='text-center'><b>Floor Price</b></p>
          </td>
          <td>
            <p><b>Discount</b></p>
          </td>
          <td>
            <p><b>Dist. Price</b></p>
          </td>
          <td>
            <p><b>Price Deviation %</b></p>
          </td>
        </tr>";
                    content += opptysubtable;
                    content += @"<!-- empty row after approval components list -->
        <tr>
          <!-- <td></td> -->
          <td colspan='5'>&nbsp;</td>
        </tr>
        <tr class='bg-grey' style='background: #cccccc!important;'>
          <td colspan='5'>
            <p class='text-left fw-bold'>Remarks</p>
          </td>
        </tr>";
                    content += @"<tr>
          <td colspan='5'>
            <p>" + Reason + "</p></td></tr>";


                }
                content += @"<!-- Note: -->
        <tr class='bg-grey' style='background: #cccccc!important;'>
          <td colspan='5'>
            <p class='fw-bold'>Note:</p>
          </td>
        </tr>
        <tr>
          <td colspan='5'>
            <p>
              To approve this, please reply to this message with just
              <span class='fw-bold'>'0K'/'Approve'/'Approved'</span> <br />
              To reject this, please reply to this message with just
              <span class='fw-bold'>'Reject'/'Rejected'.</span>
            </p>
          </td>
        </tr>
        <!-- Note ends -->
 <tr style='background: #cccccc!important;'>
          <td colspan='5'>
            &nbsp;
          </td>
        </tr>
      </table>
</br>
</br>
    </div>
  </body></html>";
                #endregion
            }
            catch (Exception ex)
            {
                content = "Exception In creating  body : " + ex.Message;
            }
            return content;
        }

        #region Code commented on 16-March-2023
        //public string getEmailBody(IOrganizationService service, string Approver, Entity order, Entity account, Entity sr, string custsegment, string billcyle, string arc, ITracingService trace, string subject)
        //{

        //    string emailbody = null;
        //    try
        //    {

        //        trace.Trace("Email Body Method Started");
        //        string TempName = null; //used for all entity reference name values
        //        string TempName2 = null;//used for Building status and billing frequency.
        //        string OrderID = string.Empty;
        //        string Productname = string.Empty;
        //        string Reason = string.Empty;
        //        string SRM = string.Empty;
        //        string accountname = string.Empty;
        //        string CANID = string.Empty;
        //        string City = string.Empty;
        //        string activationdate = string.Empty;
        //        string PreviousProduct = string.Empty;
        //        string Businesssegment = string.Empty;
        //        string Typeofrequest = string.Empty;
        //        string Caseid = string.Empty;
        //        int uptype = 0;
        //        string opptysubtable = "";
        //        string SubType = string.Empty, SubSubType = string.Empty;

        //        trace.Trace("order details Started");
        //        //Order Details
        //        if (order != null)
        //        {
        //            trace.Trace("order details are not null");
        //            //Order ID
        //            if (order.Attributes.Contains("ordernumber"))
        //            {
        //                OrderID = order.Attributes["ordernumber"].ToString();
        //            }
        //            //Product
        //            if (order.Attributes.Contains("spectra_product"))
        //            {
        //                Productname = ((EntityReference)order.Attributes["spectra_product"]).Name;
        //            }
        //            //Reason
        //            if (order.Attributes.Contains("spectra_discountreason"))
        //            {
        //                Reason = order.Attributes["spectra_discountreason"].ToString();
        //            }

        //            uptype = order.GetAttributeValue<OptionSetValue>("prioritycode").Value;

        //            trace.Trace("order details completed");
        //        }
        //        //Account Details
        //        trace.Trace("account details are started");
        //        if (account != null)
        //        {
        //            trace.Trace("account details are not null");
        //            //SRM
        //            if (account.Attributes.Contains("spectra_servicerelationshipmanagerid"))
        //            {
        //                SRM = ((EntityReference)account.Attributes["spectra_servicerelationshipmanagerid"]).Name;
        //            }
        //            //account name
        //            if (account.Attributes.Contains("name"))
        //            {
        //                accountname = account.Attributes["name"].ToString();
        //            }
        //            //CAN ID
        //            if (account.Attributes.Contains("alletech_accountid"))
        //            {
        //                CANID = account.Attributes["alletech_accountid"].ToString();
        //            }
        //            //City
        //            if (account.Attributes.Contains("alletech_city"))
        //            {
        //                City = ((EntityReference)account.Attributes["alletech_city"]).Name;
        //            }
        //            //activationdate
        //            if (account.Attributes.Contains("alletech_activationdate"))
        //            {
        //                activationdate = account.Attributes["alletech_activationdate"].ToString();
        //            }
        //            //PreviousProduct
        //            if (account.Attributes.Contains("alletech_product"))
        //            {
        //                PreviousProduct = ((EntityReference)account.Attributes["alletech_product"]).Name;
        //            }
        //            trace.Trace("account details are completed");

        //        }

        //        //Case Details
        //        if (sr != null)
        //        {
        //            trace.Trace("sr details started");
        //            //Businesssegment
        //            if (sr.Attributes.Contains("alletech_businesssegment"))
        //            {
        //                Businesssegment = ((EntityReference)sr.Attributes["alletech_businesssegment"]).Name;
        //            }
        //            //Typeofrequest Upgrade or Down grade
        //            if (sr.Attributes.Contains("alletech_subdisposition"))
        //            {
        //                Typeofrequest = ((EntityReference)sr.Attributes["alletech_subdisposition"]).Name;
        //            }
        //            //Typeofrequest Upgrade or Down grade
        //            if (sr.Attributes.Contains("alletech_subdisposition"))
        //            {
        //                Typeofrequest = ((EntityReference)sr.Attributes["alletech_subdisposition"]).Name;
        //                SubType = ((EntityReference)sr.Attributes["alletech_subdisposition"]).Name.ToString();
        //            }
        //            if (sr.Attributes.Contains("alletech_disposition"))
        //            {
        //                SubSubType = ((EntityReference)sr.Attributes["alletech_disposition"]).Name.ToString();
        //            }
        //            //Case ID
        //            if (sr.Attributes.Contains("alletech_caseidcrm"))
        //            {
        //                Caseid = sr.Attributes["alletech_caseidcrm"].ToString();
        //            }
        //            trace.Trace("SR details End");
        //        }

        //        #region Constructing HTML
        //        emailbody = "<html><body><div>";
        //        emailbody += "Hi " + Approver + ",<br/><br/>";
        //        //type
        //        if (uptype == 111260000)
        //            emailbody += "Approval of "+ Typeofrequest + " for order Id " + OrderID + ",<br/>";
        //        else if (uptype == 111260001)
        //            emailbody += "Approval of Temporary Upgradation for order " + OrderID + ",<br/>";
        //        else if (uptype == 111260002)
        //            emailbody += "Approval of B2C Downgrade for order " + OrderID + ",<br/>";

        //        emailbody += "<p style = 'color: black; font-family: ' Arial',sans-serif; font-size: 10pt;'><b> Order Details </b></p>";
        //        emailbody += "<table class='MsoNormalTable' style='width:547pt;font-family: Arial,sans-serif; font-size: 9pt; border-collapse: collapse;' border='0' cellspacing='0' cellpadding='0'><tbody>";

        //        #region Row 1
        //        //if (Oppty.Attributes.Contains("spectra_customersegmentcode"))
        //        //    TempName = GetCustomersegment(Oppty.GetAttributeValue<OptionSetValue>("spectra_customersegmentcode").Value);
        //        //else
        //        //    TempName = null;


        //        if (custsegment == null || custsegment == string.Empty)
        //        {
        //            emailbody += @"<tr style='height: 15pt;'>
        //            <td width='239' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 179pt; height: 15pt; ' colspan='2'>
        //                <p align='center' text-align: center;'><b>Business Segment</b></p>
        //            </td>
        //            <td width='131' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 98pt; height: 15pt; '>
        //                <p align='center' text-align: center;'>" + Businesssegment + @"</p>
        //            </td>
        //            <td width='119' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 89pt; height: 15pt; '>
        //                <p><b>Account Manager</b></p>
        //            </td>
        //            <td width='241' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 181pt; height: 15pt; ' colspan='2'>
        //                <p align='center' text-align: center;'>" + SRM + @"</p>
        //            </td>                   
        //        </tr>";

        //        }
        //        else
        //        {
        //            emailbody += @"<tr style ='height: 15pt;'>
        //                        <td width ='142' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 106.35pt; height: 15pt; '>
        //                        <p><b> Business Segment </b></p></td>
        //                        <td width ='97' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 72.65pt; height: 15pt; '>

        //                        <p align ='center' text-align: 'center;'>" + Businesssegment + @"</p></td>

        //                        <td width ='131' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 98pt; height: 15pt; '>
        //                        <p><b> Customer Segment </b></p></td>
        //                        <td width ='119' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 89pt; height: 15pt; '>
                                
        //                        <p align ='center' text-align: 'center;'>" + custsegment + @"</p></td>
                                
        //                        <td width ='93' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 70pt; height: 15pt; '>
        //                        <p><b> SRM Name </b></p></td>
        //                        <td width ='148' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 111pt; height: 15pt; '>
                                
        //                        <p align ='center' text-align: 'center;'>" + SRM + @"</p></td></tr>";
        //        }

        //        #endregion

        //        #region New Row added
        //        emailbody += @"<tr style='height: 15pt;'>
        //            <td width='239' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 179pt; height: 15pt; ' colspan='2'>
        //                <p align='center' text-align: center;'><b>Type</b></p>
        //            </td>
        //            <td width='131' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 98pt; height: 15pt; '>
        //                <p align='center' text-align: center;'>" + SubType + @"</p>
        //            </td>
        //            <td width='119' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 89pt; height: 15pt; '>
        //                <p><b>Sub Type</b></p>
        //            </td>
        //            <td width='241' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 181pt; height: 15pt; ' colspan='2'>
        //                <p align='center' text-align: center;'>" + SubSubType + @"</p>
        //            </td>                 
        //        </tr>";
        //        #endregion

        //        #region Row 2
        //        //if (Oppty.Attributes.Contains("alletech_companynamebusiness"))
        //        //    TempName = Oppty.GetAttributeValue<string>("alletech_companynamebusiness");
        //        //else
        //        //    TempName = null;

        //        emailbody += @"<tr style ='height: 15pt;'>
        //                        <td width ='142' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 106.35pt; height: 15pt; '>
        //                        <p><b> Customer Name </b></p></td>
        //                        <td width ='97' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 72.65pt; height: 15pt; '>

        //                        <p align ='center' text-align: 'center;'>" + accountname + @"</p></td>

        //                        <td width ='131' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 98pt; height: 15pt; '>
        //                        <p><b> Service ID </b></p></td>
        //                        <td width ='119' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 89pt; height: 15pt; '>
                                
        //                        <p align ='center' text-align: 'center;'>" + CANID + @"</p></td>
                                
        //                        <td width ='93' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 70pt; height: 15pt; '>
        //                        <p><b> City </b></p></td>
        //                        <td width ='148' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 111pt; height: 15pt; '>
                                
        //                        <p align ='center' text-align: 'center;'>" + City + @"</p></td></tr>";
        //        #endregion

        //        #region Row 3
        //        //if (Oppty.Attributes.Contains("alletech_area"))
        //        //    TempName = Oppty.GetAttributeValue<EntityReference>("alletech_area").Name;
        //        //else
        //        //    TempName = null;

        //        emailbody += @"<tr style ='height: 15pt;'>
        //                        <td width ='142' style='padding: 0in 5.4pt; border: 1pt solid windowtext; width: 106.35pt; height: 15pt; '>
        //                        <p><b> Customer Activation date </b></p></td>
        //                        <td width ='97' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 72.65pt; height: 15pt; '>

        //                        <p align ='center' text-align: 'center;'>" + activationdate + @"</p></td>

        //                        <td width ='131' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 98pt; height: 15pt; '>
        //                        <p><b> Current Plan </b></p></td>
        //                        <td width ='119' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 89pt; height: 15pt; '>
                                
        //                        <p align ='center' text-align: 'center;'>" + PreviousProduct + @"</p></td>";


        //        if (uptype == 111260000)//Permanent
        //        {
        //            emailbody += @"<td width ='93' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 70pt; height: 15pt; '>
        //                        <p><b> Current ARC </b></p></td>
        //                        <td width ='148' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 111pt; height: 15pt; '>
                                
        //                        <p align ='center' text-align: 'center;'>" + arc + @"</p></td></tr>";
        //        }
        //        else if (uptype == 111260001) //Temp
        //        {
        //            DateTime strtdate = order.GetAttributeValue<DateTime>("spectra_effectivedate");
        //            DateTime enddate = order.GetAttributeValue<DateTime>("submitdate");

        //            double days = (enddate - strtdate).TotalDays;

        //            emailbody += @"<td width ='93' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 70pt; height: 15pt; '>
        //                        <p><b> Number of days </b></p></td>
        //                        <td width ='148' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 111pt; height: 15pt; '>
                                
        //                        <p align ='center' text-align: 'center;'>" + days + @"</p></td></tr>";
        //        }
        //        else if(uptype == 111260002)//Home 
        //        {
        //            emailbody += @"<td width ='93' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 70pt; height: 15pt; '>
        //                        <p><b> Current MRC </b></p></td>
        //                        <td width ='148' style='border-width: 1pt 1pt 1pt 0px; border-style: solid solid solid none;  padding: 0in 5.4pt; width: 111pt; height: 15pt; '>
                                
        //                        <p align ='center' text-align: 'center;'>" + arc + @"</p></td></tr>";
        //        }
        //        #endregion

        //        if (uptype == 111260000)//Permanent
        //        {
        //            #region Checking OpptyProds
        //            string[] Opptyvalues = new string[4];
        //            TempName = null; TempName2 = null;

        //            string[] ordprodattr = { "productid", "priceperunit", "manualdiscountamount", "extendedamount", "productdescription" };
        //            EntityCollection ordProds = GetResultsByAttribute(service, "salesorderdetail", "salesorderid", order.Id.ToString(), ordprodattr);

        //            if (ordProds.Entities.Count > 0)
        //            {
        //                foreach (Entity ordprod in ordProds.Entities)
        //                {
        //                    string prodname = null;
        //                    Entity prod = null;
        //                    string[] prodattr = { "alletech_billingcycle", "alletech_plantype", "alletech_chargetype" };
        //                    if (ordprod.Attributes.Contains("productid"))
        //                    {
        //                        prodname = ordprod.GetAttributeValue<EntityReference>("productid").Name;
        //                        prod = GetResultByAttribute(service, "product", "productid", ordprod.GetAttributeValue<EntityReference>("productid").Id.ToString(), prodattr);
        //                    }
        //                    else
        //                    {
        //                        prodname = ordprod.GetAttributeValue<string>("productdescription");
        //                        prod = GetResultByAttribute(service, "product", "name", prodname, prodattr);
        //                    }

        //                    try
        //                    {
        //                        int plan = prod.GetAttributeValue<OptionSetValue>("alletech_plantype").Value;
        //                        int charge = prod.GetAttributeValue<OptionSetValue>("alletech_chargetype").Value;

        //                        decimal price = 0, extnamt = 0, percentAge = 0;
        //                        //if (prodname.EndsWith("RC") || prodname.EndsWith("OTC") || plan == 569480002) //addon plan type
        //                        if (charge == 569480001)// || charge == 569480002 || plan == 569480002) //addon plan type
        //                        {
        //                            price = ordprod.GetAttributeValue<Money>("priceperunit").Value;
        //                            Opptyvalues[0] = price.ToString();

        //                            if (ordprod.Contains("manualdiscountamount"))
        //                                Opptyvalues[1] = ordprod.GetAttributeValue<Money>("manualdiscountamount").Value.ToString();
        //                            else
        //                                Opptyvalues[1] = "0";

        //                            extnamt = ordprod.GetAttributeValue<Money>("extendedamount").Value;
        //                            Opptyvalues[2] = extnamt.ToString();

        //                            if (extnamt < price)
        //                            {
        //                                percentAge = (price - extnamt) / price * 100;
        //                                percentAge = decimal.Round(percentAge, 2, MidpointRounding.AwayFromZero);
        //                            }

        //                            Opptyvalues[3] = percentAge.ToString() + "%";

        //                            opptysubtable += "<tr style='height: 15pt;'>";
        //                            opptysubtable += "<td width='142' style='border-width: 0px 1pt 1pt; border-style: none solid solid; padding: 0in 5.4pt; width: 106.35pt;'> ";
        //                            opptysubtable += "<p align='center' text-align: center;'><b>" + prodname + "</b></p>";
        //                            opptysubtable += "</td>";
        //                            opptysubtable += "<td width='97' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none;  padding: 0in 5.4pt; width: 72.65pt;'> ";
        //                            opptysubtable += "<p align='center' text-align: center;'>" + Opptyvalues[0] + @"</p>";
        //                            opptysubtable += "</td>";
        //                            opptysubtable += "<td width='131' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none;  padding: 0in 5.4pt; width: 98pt;'> ";
        //                            opptysubtable += "<p align='center' text-align: center;'>" + Opptyvalues[1] + @"</p>";
        //                            opptysubtable += "</td>";
        //                            opptysubtable += "<td width='119' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none;  padding: 0in 5.4pt; width: 89pt;'> ";
        //                            opptysubtable += "<p align='center' text-align: center;'>" + Opptyvalues[2] + @"</p>";
        //                            opptysubtable += "</td>";
        //                            opptysubtable += "<td width='241' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none;  padding: 0in 5.4pt; width: 181pt;' colspan='2'> ";
        //                            opptysubtable += "<p align='center' text-align: center;'>" + Opptyvalues[3] + @"</p>";
        //                            opptysubtable += "</td>";
        //                            opptysubtable += "</tr>";
        //                        }
        //                        else
        //                        {
        //                            TempName = prodname;

        //                            if (prod.Attributes.Contains("alletech_billingcycle"))
        //                                TempName2 = prod.GetAttributeValue<EntityReference>("alletech_billingcycle").Name;
        //                        }
        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        opptysubtable += "Exception In building sub table : " + ex.Message;
        //                    }
        //                }
        //            }
        //            #endregion
        //        }

        //        #region Row 4 & 5
        //        emailbody += @"<tr style='height: 15pt;'>
        //            <td width='239' style='border-width: 0px 1pt 1pt; border-style: none solid solid; border-color: rgb(0, 0, 0) black windowtext windowtext; padding: 0in 5.4pt; width: 179pt; height: 15pt; ' colspan='2'>
        //                <p align='center' text-align: center;'><b>Requested (new) Plan</b></p>
        //            </td>
        //            <td width='131' style='border-width: 0px 0px 1pt; border-style: none none solid; border-color: rgb(0, 0, 0) rgb(0, 0, 0) windowtext; padding: 0in 5.4pt; width: 98pt; height: 15pt; '>
        //                <p align='center' text-align: center;'>" + Productname + @"</p>
        //            </td>
        //            <td width='119' style='border-width: 0px 1pt 1pt; border-style: none solid solid; padding: 0in 5.4pt; width: 89pt; height: 15pt; '>
        //                <p><b>Billing Frequency</b></p>
        //            </td>
        //            <td width='241' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none; border-color: rgb(0, 0, 0) black windowtext rgb(0, 0, 0); padding: 0in 5.4pt; width: 181pt; height: 15pt; ' colspan='2'>
        //                <p align='center' text-align: center;'>" + billcyle + @"</p>
        //            </td>
        //        </tr>";

        //        emailbody += @"<tr style='height: 15pt;'>
        //            <td width='239' style='border-width: 0px 1pt 1pt; border-style: none solid solid; border-color: rgb(0, 0, 0) black windowtext windowtext; padding: 0in 5.4pt; width: 179pt; height: 15pt; ' colspan='2'>
        //                <p align='center' text-align: center;'><b>Type of Request</b></p>
        //            </td>
        //            <td width='131' style='border-width: 0px 0px 1pt; border-style: none none solid; border-color: rgb(0, 0, 0) rgb(0, 0, 0) windowtext; padding: 0in 5.4pt; width: 98pt; height: 15pt; '>
        //                <p align='center' text-align: center;'>" + Typeofrequest + @"</p>
        //            </td>
        //            <td width='119' style='border-width: 0px 1pt 1pt; border-style: none solid solid; padding: 0in 5.4pt; width: 89pt; height: 15pt; '>
        //                <p><b>SR No.</b></p>
        //            </td>
        //            <td width='241' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none; border-color: rgb(0, 0, 0) black windowtext rgb(0, 0, 0); padding: 0in 5.4pt; width: 181pt; height: 15pt; ' colspan='2'>
        //                <p align='center' text-align: center;'>" + Caseid + @"</p>
        //            </td>
        //        </tr>";
        //        #endregion

        //        if (uptype == 111260000)//Permanent
        //        {
        //            #region Row 6
        //            emailbody += @"<tr style='height: 15pt;'>
        //            <td width='142' style='border-width: 0px 1pt 1pt; border-style: none solid solid; padding: 0in 5.4pt; width: 106.35pt; height: 15pt; '>
        //                <p align='center' text-align: center;'><b>Product Component</b></p>
        //            </td>
        //            <td width='97' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none;  padding: 0in 5.4pt; width: 72.65pt; height: 15pt; '>
        //                <p align='center' text-align: center;'><b>Floor Price</b></p>
        //            </td>
        //            <td width='131' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none;  padding: 0in 5.4pt; width: 98pt; height: 15pt; '>
        //                <p align='center' text-align: center;'><b>Discount</b></p>
        //            </td>
        //            <td width='119' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none;  padding: 0in 5.4pt; width: 89pt; height: 15pt; '>
        //                <p align='center' text-align: center;'><b>Discounted Price</b></p>
        //            </td>
        //            <td width='241' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none;  padding: 0in 5.4pt; width: 181pt; height: 15pt; ' colspan='2'>
        //                <p align='center' text-align: center;'><b>Price Deviation %</b></p>
        //            </td>
        //        </tr>";
        //            #endregion

        //            emailbody += opptysubtable;

        //            #region Row 9

        //            emailbody += @"<tr style='height: 33pt;'>
        //            <td width='142' style='border-width: 0px 1pt 1pt; border-style: none solid solid; padding: 0in 5.4pt; width: 106.35pt; height: 33pt; '>
        //                <p><b>Remarks</b></p>
        //            </td>
        //            <td width='588' style='border-width: 0px 1pt 1pt 0px; border-style: none solid solid none;  padding: 0in 5.4pt; width: 440.65pt; height: 33pt; ' colspan='5'>
        //                <p align='center' text-align: center;'>" + Reason + @"</p>
        //            </td>
        //        </tr>";
        //            #endregion
        //        }

        //        emailbody += "</tbody></table>";
        //        emailbody += "<br /> Note: To approve from email, please reply to this email with just 'Approve'/'Approved'.";
        //        emailbody += "<br />       To reject from email, please reply to this email with just 'Reject'/'Rejected'.";
        //        emailbody += "<br /><br /> Regards,<br /> Spectra Team <br />";
        //        emailbody += "</div></body></html>";

        //        #endregion
        //        //}
        //    }
        //    catch (Exception ex)
        //    {
        //        emailbody = "Exception In creating  body : " + ex.Message;
        //    }
        //    return emailbody;
        //}
        
        #endregion
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
