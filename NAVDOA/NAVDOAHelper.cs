﻿using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NAVDOA
{
    public class NAVDOAHelper
    {
        #region Commented code on 14 march 2023
        //public string getEmailBody(IOrganizationService service, Entity Img, string Approver, string Type)
        //{
        //    string emailbody = null;
        //    try
        //    {
        //        string TempName = null; //used for all entity reference name values
        //        EntityCollection TempColl = null;

        //        string customer = null;
        //        string canid = null;
        //        string seg = null;
        //        string city = null;
        //        string BSType = null;
        //        string BSName = null;
        //        string BSStatus = null;

        //        if (Type == "WCR")
        //        {
        //            seg = "Home";
        //            canid = Img.GetAttributeValue<string>("spectra_canno");
        //            customer = Img.GetAttributeValue<EntityReference>("alletech_account").Name;
        //            city = Img.GetAttributeValue<EntityReference>("alletech_city").Name;
        //            BSType = "Society";

        //            if (Img.Attributes.Contains("spectra_society"))
        //            {
        //                BSName = Img.GetAttributeValue<EntityReference>("spectra_society").Name;

        //                Entity building = GetResultByAttribute(service, "alletech_society", "alletech_societyid", Img.GetAttributeValue<EntityReference>("spectra_society").Id.ToString(), "alletech_societybuildingstatus");
        //                if (building != null && building.Attributes.Contains("alletech_societybuildingstatus"))
        //                {
        //                    BSStatus = getBuildingStatusValue(building.GetAttributeValue<OptionSetValue>("alletech_societybuildingstatus").Value);
        //                }
        //            }
        //        }
        //        else
        //        {
        //            seg = "Business";
        //            canid = Img.GetAttributeValue<string>("spectra_canid");
        //            customer = Img.GetAttributeValue<EntityReference>("alletech_companyname").Name;
        //            city = Img.GetAttributeValue<EntityReference>("alletech_city3").Name;
        //            BSType = "Building";

        //            if (Img.Attributes.Contains("alletech_building1"))
        //            {
        //                BSName = Img.GetAttributeValue<EntityReference>("alletech_building1").Name;
        //                Entity building = GetResultByAttribute(service, "alletech_building", "alletech_buildingid", Img.GetAttributeValue<EntityReference>("alletech_building1").Id.ToString(), "alletech_buildingstatus");
        //                if (building != null && building.Attributes.Contains("alletech_buildingstatus"))
        //                {
        //                    BSStatus = getBuildingStatusValue(building.GetAttributeValue<OptionSetValue>("alletech_buildingstatus").Value);
        //                }
        //            }
        //        }

        //        string fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                          <entity name='alletech_caf'>
        //                            <attribute name='alletech_product'/>
        //                            <attribute name='alletech_oppurtunityid' />
        //                            <filter type='and'>
        //                              <condition attribute='alletech_cafid' operator='eq' value='" + Img.GetAttributeValue<EntityReference>("alletech_cafid").Id + @"' />
        //                            </filter>
        //                            <link-entity name='opportunity' from='opportunityid' to='alletech_oppurtunityid' visible='false' link-type='outer' alias='opp'>
        //                              <attribute name='alletech_redundancyrequired' />
        //                              <attribute name='spectra_customersegmentcode' />
        //                            </link-entity>
        //                          </entity>
        //                        </fetch>";
        //        TempColl = service.RetrieveMultiple(new FetchExpression(fetch));

        //        Entity Caf = TempColl.Entities[0];//GetResultByAttribute(service, "opportunity", "opportunityid", Img.GetAttributeValue<EntityReference>("alletech_opportunity").Id.ToString(), new string[] { "alletech_redundancyrequired", "spectra_customersegmentcode" });

        //        #region Constructing HTML

        //        #region Headers
        //        emailbody = "<html><body><div>";
        //        emailbody += "Hi " + Approver + ",<br/><br/>";
        //        emailbody += "<p style = 'color: black; font-family: ' Arial',sans-serif; font-size: 10pt;'><b> Approval of item quantity deviation for CAN : " + canid + " </b></p>";
        //        emailbody += "<p style = 'color: black; font-family: ' Arial',sans-serif; font-size: 10pt;'><b> Order Details </b></p>";
        //        emailbody += @"<table class='MsoNormalTable' style='width:600pt;font-family: Arial,sans-serif; font-size: 9pt; border-collapse:collapse; padding: 0in 5.4pt;' cellspacing='0' cellpadding='5.4pt'>
        //                       <col width='100'><col width='90'><col width='110'><col width='100'><col width='100'><col width='100'>
        //                       <tbody>";
        //        #endregion

        //        #region Row 1

        //        emailbody += @"<tr style='height : 20pt'>
        //                       <td style='border-width: 1pt; border-style:solid;'><p><b>Business Segment</b></p></td>
        //                       <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align='center;'>" + seg + @"</p></td>";

        //        if (Caf.Attributes.Contains("opp.spectra_customersegmentcode"))
        //            TempName = GetCustomersegment(((OptionSetValue)Caf.GetAttributeValue<AliasedValue>("opp.spectra_customersegmentcode").Value).Value);
        //        else
        //            TempName = null;

        //        emailbody += @"<td style='border-width: 1pt; border-style:solid;'><p><b>Customer Segment</b></p></td>
        //                       <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align='center;'>" + TempName + @"</p></td>";

        //        emailbody += @"<td style='border-width: 1pt; border-style:solid;'><p><b>CAN ID</b></p></td>
        //                       <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align='center;'>" + canid + @"</p></td>
        //                       </tr>";
        //        #endregion

        //        #region Row 2

        //        emailbody += @"<tr style='height : 20pt'>
        //                           <td style='border-width: 1pt; border-style:solid;'><p><b>Customer Name</b></p></td>
        //                           <td style='border-width: 1pt; border-style:solid;' colspan='2'><p align='center' text-align='center;'>" + customer + @"</p></td>";

        //        if (Img.Attributes.Contains("ownerid"))
        //            TempName = Img.GetAttributeValue<EntityReference>("ownerid").Name;
        //        else
        //            TempName = null;

        //        emailbody += @"<td style='border-width: 1pt; border-style:solid;'><p><b>Vendor Name</b></p></td>
        //                           <td style='border-width: 1pt; border-style:solid;' colspan='2'><p align='center' text-align='center;'>" + TempName + @"</p></td>
        //                           </tr>";
        //        #endregion

        //        #region Row 3

        //        emailbody += @"<tr style='height : 20pt'>
        //                            <td style='border-width: 1pt; border-style:solid;'><p><b>City</b></p></td>
        //                            <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align='center;'>" + city + @"</p></td>";

        //        if (Img.Attributes.Contains("alletech_area"))
        //            TempName = Img.GetAttributeValue<EntityReference>("alletech_area").Name;
        //        else
        //            TempName = null;

        //        emailbody += @"<td style='border-width: 1pt; border-style:solid;'><b>Area</b></td>
        //                            <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align='center;'>" + TempName + @"</p></td>";

        //        emailbody += @"<td style='border-width: 1pt; border-style:solid;'><b>" + BSType + @" Name</b></td>
        //                            <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align='center;'>" + BSName + @"</p></td>
        //                            </tr>";
        //        #endregion

        //        #region Row 4
        //        emailbody += @" <tr style='height : 20pt'>
        //                            <td style='border-width: 1pt; border-style:solid;'> <b>" + BSType + @" Status</b></td>
        //                            <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align='center;'>" + BSStatus + @"</p></td>";

        //        if (Caf.Attributes.Contains("opp.alletech_redundancyrequired") && (bool)Caf.GetAttributeValue<AliasedValue>("opp.alletech_redundancyrequired").Value)
        //            TempName = "Yes";
        //        else
        //            TempName = "No";

        //        emailbody += @" <td style='border-width: 1pt; border-style:solid;'><b>Redundancy required</b></td>
        //                            <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align='center;'>" + TempName + @"</p></td>";

        //        //if (BSStatus == "Non-RFS")
        //        //{
        //        //    decimal amt = 0;
        //        //    string[] feasiblityattr = { "alletech_feasibilityid" };
        //        //    EntityCollection feasibleColl = GetResultsByAttribute(service, "alletech_feasibility", "alletech_opportunity", Caf.GetAttributeValue<EntityReference>("alletech_oppurtunityid").Id.ToString(), feasiblityattr);

        //        //    if (feasibleColl.Entities.Count > 0)
        //        //    {
        //        //        foreach (Entity feasible in feasibleColl.Entities)
        //        //        {
        //        //            string[] Prjattr = { "alletech_pjcid", "alletech_totalcalculatedcost" };
        //        //            EntityCollection ProjcstColl = GetResultsByAttribute(service, "alletech_projectestimationcost", "alletech_pjcid", feasible.Id.ToString(), Prjattr);
        //        //            if (ProjcstColl.Entities.Count > 0)
        //        //            {
        //        //                foreach (Entity prj in ProjcstColl.Entities)
        //        //                {
        //        //                    if (prj.Attributes.Contains("alletech_totalcalculatedcost"))
        //        //                        amt += prj.GetAttributeValue<Money>("alletech_totalcalculatedcost").Value;
        //        //                }
        //        //            }
        //        //            if (TempName == "No")
        //        //                break;
        //        //        }
        //        //    }
        //        //    TempName = amt.ToString();
        //        //}
        //        //else
        //        TempName = null;

        //        emailbody += @" <td style='border-width: 1pt; border-style:solid;'><b>Connectivity Cost</b></td>
        //                            <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align='center;'>" + TempName + @"</p></td>
        //                            </tr>";
        //        #endregion

        //        #region Row 5

        //        if (Caf.Attributes.Contains("alletech_product"))
        //            TempName = Caf.GetAttributeValue<EntityReference>("alletech_product").Name;
        //        else
        //            TempName = null;

        //        emailbody += @"<tr style='height : 20pt'>
        //                            <td style='border-width: 1pt; border-style:solid;'><b>Product</b></td>
        //                            <td style='border-width: 1pt; border-style:solid;' colspan='2'><p align='center' text-align='center;'>" + TempName + @"</p></td>";

        //        if (Img.Attributes.Contains("spectra_zone"))
        //            TempName = Img.GetAttributeValue<EntityReference>("spectra_zone").Name;
        //        else
        //            TempName = null;

        //        emailbody += @"<td style='border-width: 1pt; border-style:solid;'><b>Zone</b></td>
        //                            <td style='border-width: 1pt; border-style:solid;' colspan='2'><p align='center' text-align='center;'>" + TempName + @"</p></td>
        //                            </tr>";
        //        #endregion

        //        #region Row 6
        //        emailbody += @"<tr style='height : 20pt'>
        //                            <td style='border-width: 1pt; border-style:solid;' colspan='2'><p align='center' text-align=' center;'><b>Material Name</b></p></td>
        //                            <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'><b>Item Code</b></p></td>
        //                            <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'><b>Expected Qty</b></p></td>
        //                            <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'><b>Consumed Qty</b></p></td>
        //                            <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'><b>Deviation</b></p></td>
        //                            </tr>";
        //        #endregion

        //        #region Checking Item consumptions: Old Code commented on 06-June-2021

        //        //TempName = null;
        //        //string Itemtable = "";
        //        //fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //        //                  <entity name='alletech_itemconsumption'>
        //        //                    <attribute name='alletech_quantity' />
        //        //                    <attribute name='alletech_actualquantityused' />
        //        //                    <attribute name='alletech_subitem' />
        //        //                    <attribute name='spectra_itemtype' />
        //        //                    <order attribute='alletech_subitem' descending='false' />
        //        //                    <order attribute='spectra_itemtype' descending='false' />
        //        //                    <filter type='and'>";
        //        //if (Type == "WCR")
        //        //{
        //        //    fetch += "<condition attribute='alletech_wcr' operator='eq' uiname='' uitype='alletech_wcr' value='" + Img.Id.ToString() + @"' />";
        //        //}
        //        //else
        //        //{
        //        //    fetch += "<condition attribute='spectra_installationreport' operator='eq' uiname='' uitype='alletech_installationform' value='" + Img.Id.ToString() + @"' />";
        //        //}

        //        //fetch += @"       </filter>
        //        //                    <link-entity name='alletech_subitem' from='alletech_subitemid' to='alletech_subitem' visible='false' link-type='outer' alias='sub'>
        //        //                      <attribute name='alletech_subitemcode' />
        //        //                    </link-entity>
        //        //                  </entity>
        //        //                </fetch>";

        //        //TempColl = service.RetrieveMultiple(new FetchExpression(fetch));

        //        //if (TempColl.Entities.Count > 0)
        //        //{
        //        //    int count = TempColl.Entities.Count;
        //        //    for (int i = 0; i < count; i++)
        //        //    {
        //        //        try
        //        //        {
        //        //            //if additional
        //        //            if (TempColl.Entities[i].GetAttributeValue<OptionSetValue>("spectra_itemtype").Value == 111260001)
        //        //            {
        //        //                int qty = 0, conqty = 0, dev = 0;
        //        //                string item = TempColl.Entities[i].GetAttributeValue<EntityReference>("alletech_subitem").Name;
        //        //                string code = null;

        //        //                if (TempColl.Entities[i].Attributes.Contains("sub.alletech_subitemcode"))
        //        //                    code = (string)TempColl.Entities[i].GetAttributeValue<AliasedValue>("sub.alletech_subitemcode").Value;

        //        //                if (i + 1 < count && item == TempColl.Entities[i + 1].GetAttributeValue<EntityReference>("alletech_subitem").Name)
        //        //                {
        //        //                    qty = TempColl.Entities[i + 1].GetAttributeValue<int>("alletech_actualquantityused");
        //        //                }
        //        //                else
        //        //                {
        //        //                    qty = 0;
        //        //                }

        //        //                conqty = qty + TempColl.Entities[i].GetAttributeValue<int>("alletech_actualquantityused");
        //        //                dev = TempColl.Entities[i].GetAttributeValue<int>("alletech_actualquantityused");

        //        //                Itemtable += @"<tr style='height : 20pt'>
        //        //                           <td style='border-width: 1pt; border-style:solid;' colspan='2'><p align='center' text-align='center;'><b>" + item + @"</b></p></td>
        //        //                           <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'>" + code + @"</p></td>
        //        //                           <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'>" + qty + @"</p></td>
        //        //                           <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'>" + conqty + @"</p></td>
        //        //                           <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'>" + dev + @"</p></td>
        //        //                           </tr>";
        //        //            }
        //        //        }
        //        //        catch (Exception ex)
        //        //        {

        //        //        }
        //        //    }
        //        //}
        //        #endregion

        //        #region Checking Item consumptions: New Code created on 06-June-2021 by VLabs
        //        TempName = null;
        //        string Itemtable = "";
        //        fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                          <entity name='alletech_itemconsumption'>
        //                            <attribute name='alletech_quantity' />
        //                            <attribute name='alletech_actualquantityused' />
        //                            <attribute name='alletech_subitem' />
        //                            <attribute name='spectra_itemtype' />
        //                            <order attribute='alletech_subitem' descending='false' />
        //                            <filter type='and'>";
        //        if (Type == "WCR")
        //        {
        //            fetch += "<condition attribute='alletech_wcr' operator='eq' uiname='' uitype='alletech_wcr' value='" + Img.Id.ToString() + @"' />";
        //        }
        //        else
        //        {
        //            fetch += "<condition attribute='spectra_installationreport' operator='eq' uiname='' uitype='alletech_installationform' value='" + Img.Id.ToString() + @"' />";
        //        }

        //        fetch += @"       <condition attribute='spectra_itemtype' operator='eq' value='111260001' />
        //                            </filter>
        //                            <link-entity name='alletech_subitem' from='alletech_subitemid' to='alletech_subitem' visible='false' link-type='outer' alias='sub'>
        //                              <attribute name='alletech_itemcodeinnav' />
        //                            </link-entity>
        //                          </entity>
        //                        </fetch>";

        //        TempColl = service.RetrieveMultiple(new FetchExpression(fetch));

        //        if (TempColl.Entities.Count > 0)
        //        {
        //            string _duplicate_Item_Name = string.Empty;
        //            int count = TempColl.Entities.Count;
        //            for (int i = 0; i < count; i++)
        //            {
        //                try
        //                {
        //                    #region New Code
        //                    if (TempColl.Entities[i].GetAttributeValue<OptionSetValue>("spectra_itemtype").Value == 111260001)
        //                    {
        //                        int eQty = 0, additional = 0, consumedQty = 0, deviation = 0;
        //                        string ItemCode = null;
        //                        string item = TempColl.Entities[i].GetAttributeValue<EntityReference>("alletech_subitem").Name;

        //                        if (_duplicate_Item_Name != item)
        //                        {
        //                            if (TempColl.Entities[i].Attributes.Contains("sub.alletech_itemcodeinnav"))
        //                                ItemCode = (string)TempColl.Entities[i].GetAttributeValue<AliasedValue>("sub.alletech_itemcodeinnav").Value;

        //                            string ItemCodeCheck = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                                              <entity name='alletech_itemconsumption'>
        //                                                <attribute name='createdon' />
        //                                                <attribute name='alletech_actualquantityused' />
        //                                                <attribute name='spectra_itemtype' />
        //                                                <attribute name='alletech_subitem' />
        //                                                <attribute name='alletech_itemconsumptionid' />
        //                                                <order attribute='spectra_itemtype' descending='false' />
        //                                                <filter type='and'>";
        //                            if (Type == "WCR")
        //                            {
        //                                ItemCodeCheck += "<condition attribute='alletech_wcr' operator='eq' uiname='' uitype='alletech_wcr' value='" + Img.Id.ToString() + @"' />";
        //                            }
        //                            else
        //                            {
        //                                ItemCodeCheck += "<condition attribute='spectra_installationreport' operator='eq' uiname='' uitype='alletech_installationform' value='" + Img.Id.ToString() + @"' />";
        //                            }

        //                            ItemCodeCheck += @" </filter>
        //                                                <link-entity name='alletech_subitem' from='alletech_subitemid' to='alletech_subitem' alias='ac'>
        //                                                  <filter type='and'>
        //                                                    <condition attribute='alletech_itemcodeinnav' operator='eq' value='" + ItemCode + @"' />
        //                                                  </filter>
        //                                                </link-entity>
        //                                              </entity>
        //                                            </fetch>";
        //                            EntityCollection ItemCodeColl = service.RetrieveMultiple(new FetchExpression(ItemCodeCheck));
        //                            if (ItemCodeColl.Entities.Count > 0)
        //                            {
        //                                foreach (Entity _default in ItemCodeColl.Entities)
        //                                {
        //                                    if (_default.Attributes.Contains("spectra_itemtype"))
        //                                    {
        //                                        if (_default.GetAttributeValue<OptionSetValue>("spectra_itemtype").Value == 111260000)
        //                                        {
        //                                            eQty += _default.GetAttributeValue<int>("alletech_actualquantityused");
        //                                            //consumedQty += eQty;
        //                                        }
        //                                    }
        //                                }
        //                                foreach (Entity _additional in ItemCodeColl.Entities)
        //                                {
        //                                    if (_additional.Attributes.Contains("spectra_itemtype"))
        //                                    {
        //                                        if (_additional.GetAttributeValue<OptionSetValue>("spectra_itemtype").Value == 111260001)
        //                                        {
        //                                            additional += _additional.GetAttributeValue<int>("alletech_actualquantityused");
        //                                            //consumedQty += eQty;
        //                                        }
        //                                    }
        //                                }

        //                            }
        //                            _duplicate_Item_Name = TempColl.Entities[i].GetAttributeValue<EntityReference>("alletech_subitem").Name;
        //                            consumedQty = eQty + additional;
        //                            deviation = consumedQty - eQty;
        //                            Itemtable += @"<tr style='height : 20pt'>
        //                                               <td style='border-width: 1pt; border-style:solid;' colspan='2'><p align='center' text-align='center;'><b>" + item + @"</b></p></td>
        //                                               <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'>" + ItemCode + @"</p></td>
        //                                               <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'>" + eQty + @"</p></td>
        //                                               <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'>" + consumedQty + @"</p></td>
        //                                               <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'>" + deviation + @"</p></td>
        //                                               </tr>";

        //                        }
        //                    }
        //                    #endregion
        //                }
        //                catch (Exception ex)
        //                {
        //                    throw new InvalidPluginExecutionException("Error in Item Consumption construct: " + ex.Message);
        //                }
        //            }
        //        }

        //        #endregion

        //        #region Checking Installation Item:  Old Code commented on 07-June-2021 byVLabs
        //        //if (Type == "IR")
        //        //{
        //        //    TempName = null;

        //        //    fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //        //                  <entity name='alletech_installationitem'>
        //        //                    <attribute name='spectra_quantityir' />
        //        //                    <attribute name='alletech_subitem' />
        //        //                    <attribute name='spectra_itemtype' />
        //        //                    <order attribute='alletech_subitem' descending='false' />
        //        //                    <order attribute='spectra_itemtype' descending='false' />
        //        //                    <filter type='and'>
        //        //                      <condition attribute='alletech_installationform' operator='eq' uiname='' uitype='alletech_installationform' value='" + Img.Id.ToString() + @"' />
        //        //                    </filter>
        //        //                    <link-entity name='alletech_subitem' from='alletech_subitemid' to='alletech_subitem' visible='false' link-type='outer' alias='sub'>
        //        //                      <attribute name='alletech_subitemcode' />
        //        //                    </link-entity>
        //        //                  </entity>
        //        //                </fetch>";

        //        //    TempColl = service.RetrieveMultiple(new FetchExpression(fetch));

        //        //    if (TempColl.Entities.Count > 0)
        //        //    {
        //        //        int count = TempColl.Entities.Count;
        //        //        for (int i = 0; i < count; i++)
        //        //        {
        //        //            try
        //        //            {
        //        //                //if additional
        //        //                if (TempColl.Entities[i].GetAttributeValue<OptionSetValue>("spectra_itemtype").Value == 111260001)
        //        //                {
        //        //                    int qty = 0, conqty = 0, dev = 0;
        //        //                    string item = TempColl.Entities[i].GetAttributeValue<EntityReference>("alletech_subitem").Name;
        //        //                    string code = null;

        //        //                    if (TempColl.Entities[i].Attributes.Contains("sub.alletech_subitemcode"))
        //        //                        code = (string)TempColl.Entities[i].GetAttributeValue<AliasedValue>("sub.alletech_subitemcode").Value;

        //        //                    if (i + 1 < count && item == TempColl.Entities[i + 1].GetAttributeValue<EntityReference>("alletech_subitem").Name)
        //        //                    {
        //        //                        qty = TempColl.Entities[i + 1].GetAttributeValue<int>("spectra_quantityir");
        //        //                    }
        //        //                    else
        //        //                    {
        //        //                        qty = 0;
        //        //                    }

        //        //                    conqty = qty + TempColl.Entities[i].GetAttributeValue<int>("spectra_quantityir");
        //        //                    dev = TempColl.Entities[i].GetAttributeValue<int>("spectra_quantityir");

        //        //                    Itemtable += @"<tr style='height : 20pt'>
        //        //                           <td style='border-width: 1pt; border-style:solid;' colspan='2'><p align='center' text-align='center;'><b>" + item + @"</b></p></td>
        //        //                           <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'>" + code + @"</p></td>
        //        //                           <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'>" + qty + @"</p></td>
        //        //                           <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'>" + conqty + @"</p></td>
        //        //                           <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'>" + dev + @"</p></td>
        //        //                           </tr>";
        //        //                }
        //        //            }
        //        //            catch (Exception ex)
        //        //            {
        //        //                throw new InvalidPluginExecutionException("Error in Installation item construction: " + ex.Message);
        //        //            }
        //        //        }
        //        //    }
        //        //}

        //        #endregion

        //        #region Checking Installation Item: New Code on 07-June-2021 by VLabs
        //        if (Type == "IR")
        //        {
        //            TempName = null;
        //            fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                      <entity name='alletech_installationitem'>
        //                        <attribute name='createdon' />
        //                        <attribute name='spectra_itemtype' />
        //                        <attribute name='alletech_installationitemid' />
        //                        <attribute name='alletech_subitem' />
        //                        <attribute name='spectra_quantityir' />
        //                        <order attribute='alletech_subitem' descending='false' />
        //                        <filter type='and'>
        //                          <condition attribute='spectra_itemtype' operator='eq' value='111260001' />
        //                          <condition attribute='alletech_installationform' operator='eq' value='" + Img.Id.ToString() + @"' />
        //                        </filter>
        //                        <link-entity name='alletech_subitem' from='alletech_subitemid' to='alletech_subitem' visible='false' link-type='outer' alias='sub'>
        //                          <attribute name='alletech_itemcodeinnav' />
        //                        </link-entity>
        //                      </entity>
        //                    </fetch>";
        //            TempColl = service.RetrieveMultiple(new FetchExpression(fetch));
        //            if (TempColl.Entities.Count > 0)
        //            {
        //                string _duplicate_InstallationItem_Name = string.Empty;
        //                int count = TempColl.Entities.Count;
        //                for (int i = 0; i < count; i++)
        //                {
        //                    try
        //                    {
        //                        //if additional
        //                        if (TempColl.Entities[i].GetAttributeValue<OptionSetValue>("spectra_itemtype").Value == 111260001)
        //                        {
        //                            int eQty = 0, additional = 0, consumedQty = 0, deviation = 0;
        //                            string ItemCode = null;
        //                            string item = TempColl.Entities[i].GetAttributeValue<EntityReference>("alletech_subitem").Name;
        //                            if (_duplicate_InstallationItem_Name != item)
        //                            {
        //                                if (TempColl.Entities[i].Attributes.Contains("sub.alletech_itemcodeinnav"))
        //                                    ItemCode = (string)TempColl.Entities[i].GetAttributeValue<AliasedValue>("sub.alletech_itemcodeinnav").Value;
        //                                string _installationFetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                                                              <entity name='alletech_installationitem'>
        //                                                                <attribute name='createdon' />
        //                                                                <attribute name='spectra_itemtype' />
        //                                                                <attribute name='alletech_installationitemid' />
        //                                                                <attribute name='alletech_subitem' />
        //                                                                <attribute name='spectra_quantityir' />
        //                                                                <order attribute='spectra_itemtype' descending='false' />
        //                                                                <filter type='and'>
        //                                                                  <condition attribute='alletech_installationform' operator='eq' value='" + Img.Id.ToString() + @"' />
        //                                                                </filter>
        //                                                                <link-entity name='alletech_subitem' from='alletech_subitemid' to='alletech_subitem' alias='sub'>
        //                                                                  <attribute name='alletech_itemcodeinnav' />
        //                                                                  <filter type='and'>
        //                                                                    <condition attribute='alletech_itemcodeinnav' operator='eq' value='" + ItemCode + @"' />
        //                                                                  </filter>
        //                                                                </link-entity>
        //                                                              </entity>
        //                                                            </fetch>";
        //                                EntityCollection _installationItemCodeColl = service.RetrieveMultiple(new FetchExpression(_installationFetch));
        //                                if (_installationItemCodeColl.Entities.Count > 0)
        //                                {
        //                                    foreach (Entity _default in _installationItemCodeColl.Entities)
        //                                    {
        //                                        if (_default.Attributes.Contains("spectra_itemtype"))
        //                                        {
        //                                            if (_default.GetAttributeValue<OptionSetValue>("spectra_itemtype").Value == 111260000)
        //                                            {
        //                                                eQty += _default.GetAttributeValue<int>("spectra_quantityir");
        //                                            }
        //                                        }
        //                                    }
        //                                    foreach (Entity _additional in _installationItemCodeColl.Entities)
        //                                    {
        //                                        if (_additional.Attributes.Contains("spectra_itemtype"))
        //                                        {
        //                                            if (_additional.GetAttributeValue<OptionSetValue>("spectra_itemtype").Value == 111260001)
        //                                            {
        //                                                additional += _additional.GetAttributeValue<int>("spectra_quantityir");
        //                                            }
        //                                        }

        //                                        //if (_itemCode.Attributes.Contains("spectra_itemtype"))
        //                                        //{
        //                                        //    if (_itemCode.GetAttributeValue<OptionSetValue>("spectra_itemtype").Value == 111260000)
        //                                        //    {
        //                                        //        eQty += _itemCode.GetAttributeValue<int>("spectra_quantityir");
        //                                        //        consumedQty += eQty;
        //                                        //    }
        //                                        //    else
        //                                        //    {
        //                                        //        eQty = 0;
        //                                        //        consumedQty += _itemCode.GetAttributeValue<int>("spectra_quantityir");
        //                                        //    }
        //                                        //}
        //                                    }
        //                                    _duplicate_InstallationItem_Name = TempColl.Entities[i].GetAttributeValue<EntityReference>("alletech_subitem").Name;
        //                                    consumedQty = eQty + additional;
        //                                    deviation = consumedQty - eQty;
        //                                    Itemtable += @"<tr style='height : 20pt'>
        //                                               <td style='border-width: 1pt; border-style:solid;' colspan='2'><p align='center' text-align='center;'><b>" + item + @"</b></p></td>
        //                                               <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'>" + ItemCode + @"</p></td>
        //                                               <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'>" + eQty + @"</p></td>
        //                                               <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'>" + consumedQty + @"</p></td>
        //                                               <td style='border-width: 1pt; border-style:solid;'><p align='center' text-align=' center;'>" + deviation + @"</p></td>
        //                                               </tr>";
        //                                }
        //                            }
        //                        }
        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        throw new InvalidPluginExecutionException("Error in Installation Item Quantity adding in DOA: " + ex.Message);
        //                    }
        //                }
        //            }
        //        }
        //        #endregion

        //        emailbody += Itemtable;//Items

        //        #region Remarks Row

        //        if (Img.Attributes.Contains("alletech_remark"))
        //            TempName = Img.GetAttributeValue<string>("alletech_remark");
        //        else
        //            TempName = null;

        //        emailbody += @"<tr style='height: 25pt;'>
        //                            <td style='border-width: 1pt; border-style:solid;'><p><b>Remarks</b></p></td>
        //                            <td style='border-width: 1pt; border-style:solid;' colspan='5'><p align='center' text-align=' center;'>" + TempName + @"</p></td>
        //                            </tr>";
        //        #endregion

        //        emailbody += "</tbody></table>";
        //        emailbody += "<br /> Note: To approve from email, please reply to this email with just 'Approve'/'Approved'.";
        //        emailbody += "<br />       To reject from email, please reply to this email with just 'Reject'/'Rejected'.";
        //        emailbody += "<br /><br /> Regards,<br /> Spectra Team <br />";
        //        emailbody += "</div></body></html>";

        //        #endregion

        //    }
        //    catch (Exception ex)
        //    {
        //        emailbody = "Exception In creating  body : " + ex.Message;
        //    }
        //    return emailbody;
        //}
        #endregion

        public string getEmailBody(IOrganizationService service, Entity Img, string Approver, string Type, string subject)
        {
            string TempName = string.Empty;
            string content = string.Empty;
            try
            {
                EntityCollection TempColl = null;

                string customer = null;
                string canid = null;
                string seg = null;
                string city = null;
                string BSType = null;
                string BSName = null;
                string BSStatus = null;
                string BSStatusName = null;

                if (Type == "WCR")
                {
                    seg = "Home";
                    canid = Img.GetAttributeValue<string>("spectra_canno");
                    customer = Img.GetAttributeValue<EntityReference>("alletech_account").Name;
                    city = Img.GetAttributeValue<EntityReference>("alletech_city").Name;
                    BSType = "Society Name";
                    BSStatusName = "Society Status";

                    if (Img.Attributes.Contains("spectra_society"))
                    {
                        BSName = Img.GetAttributeValue<EntityReference>("spectra_society").Name;

                        Entity building = GetResultByAttribute(service, "alletech_society", "alletech_societyid", Img.GetAttributeValue<EntityReference>("spectra_society").Id.ToString(), "alletech_societybuildingstatus");
                        if (building != null && building.Attributes.Contains("alletech_societybuildingstatus"))
                        {
                            BSStatus = getBuildingStatusValue(building.GetAttributeValue<OptionSetValue>("alletech_societybuildingstatus").Value);
                        }
                    }
                }
                else
                {
                    seg = "Business";
                    canid = Img.GetAttributeValue<string>("spectra_canid");
                    customer = Img.GetAttributeValue<EntityReference>("alletech_companyname").Name;
                    city = Img.GetAttributeValue<EntityReference>("alletech_city3").Name;
                    BSType = "Building";
                    BSStatusName = "Building Status";

                    if (Img.Attributes.Contains("alletech_building1"))
                    {
                        BSName = Img.GetAttributeValue<EntityReference>("alletech_building1").Name;
                        Entity building = GetResultByAttribute(service, "alletech_building", "alletech_buildingid", Img.GetAttributeValue<EntityReference>("alletech_building1").Id.ToString(), "alletech_buildingstatus");
                        if (building != null && building.Attributes.Contains("alletech_buildingstatus"))
                        {
                            BSStatus = getBuildingStatusValue(building.GetAttributeValue<OptionSetValue>("alletech_buildingstatus").Value);
                        }
                    }
                }

                string fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='alletech_caf'>
                                    <attribute name='alletech_product'/>
                                    <attribute name='alletech_oppurtunityid' />
                                    <filter type='and'>
                                      <condition attribute='alletech_cafid' operator='eq' value='" + Img.GetAttributeValue<EntityReference>("alletech_cafid").Id + @"' />
                                    </filter>
                                    <link-entity name='opportunity' from='opportunityid' to='alletech_oppurtunityid' visible='false' link-type='outer' alias='opp'>
                                      <attribute name='alletech_redundancyrequired' />
                                      <attribute name='spectra_customersegmentcode' />
                                    </link-entity>
                                  </entity>
                                </fetch>";
                TempColl = service.RetrieveMultiple(new FetchExpression(fetch));

                Entity Caf = TempColl.Entities[0];//GetResultByAttribute(service, "opportunity", "opportunityid", Img.GetAttributeValue<EntityReference>("alletech_opportunity").Id.ToString(), new string[] { "alletech_redundancyrequired", "spectra_customersegmentcode" });

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
  background-color: #cccccc !importan;
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
                content += "<p>" + subject + "</p> </td></tr>";
                content += "<!-- Details --><!-- Order ID -->";
                content += "<tr style='background: #cccccc!important;'><td colspan='2'><p class='text-left'>Order Id:</p></td><td colspan='3'><p class='text-left'>" + canid + "</p></td></tr>";
                content += "<!-- Segment -->";
                if (Caf.Attributes.Contains("opp.spectra_customersegmentcode"))
                    TempName = GetCustomersegment(((OptionSetValue)Caf.GetAttributeValue<AliasedValue>("opp.spectra_customersegmentcode").Value).Value);
                else
                    TempName = null;

                content += "<tr><td colspan='2'><p class='text-left'>Segment</p></td><td colspan='3'><p class='text-left'>" + seg + "</p></td></tr>";
                content += "<!-- Customer Name -->";
                content += @"<tr style='background: #cccccc!important;'>
          <td colspan='2'>
            <p class='text-left'>Customer Name</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + customer + "</p></td></tr>";
                content += "<!-- Vendor Name -->";
                if (Img.Attributes.Contains("ownerid"))
                    TempName = Img.GetAttributeValue<EntityReference>("ownerid").Name;
                else
                    TempName = null;

                content += @"<tr>
          <td colspan='2'>
            <p class='text-left'>Vendor Name</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + TempName + "</p></td></tr>";
                content += "<!-- city -->";
                content += @"<tr style='background: #cccccc!important;'>
          <td colspan='2'>
            <p class='text-left'>City</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + city + "</p></td></tr>";
                content += "<!-- Area -->";
                if (Img.Attributes.Contains("alletech_area"))
                    TempName = Img.GetAttributeValue<EntityReference>("alletech_area").Name;
                else
                    TempName = null;

                content += @"<tr>
          <td colspan='2'>
            <p class='text-left'>Area</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + TempName + "</p></td></tr>";
                content += @"<tr style='background: #cccccc!important;'>
          <td colspan='2'>
            <p class='text-left'>" + BSType + "</p></td><td colspan='3'><p class='text-left'>" + BSName + "</p></td></tr>";
                content += @"<tr>
          <td colspan='2'>
            <p class='text-left'>" + BSStatusName + "</p></td><td colspan='3'><p class='text-left'>" + BSStatus + "</p></td></tr>";
                content += "<!-- Redundancy required -->";
                if (Caf.Attributes.Contains("opp.alletech_redundancyrequired") && (bool)Caf.GetAttributeValue<AliasedValue>("opp.alletech_redundancyrequired").Value)
                    TempName = "Yes";
                else
                    TempName = "No";
                content += @"<tr style='background: #cccccc!important;'>
          <td colspan='2'>
            <p class='text-left'>Redundancy required</p>
          </td>
         

          <td colspan='3'>
            <p class='text-left'>" + TempName + "</p></td></tr>";
                content += "<!-- Product -->";
                if (Caf.Attributes.Contains("alletech_product"))
                    TempName = Caf.GetAttributeValue<EntityReference>("alletech_product").Name;
                else
                    TempName = null;
                content += @"<tr>
          <td colspan='2'>
            <p class='text-left'>Product</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + TempName + "</p></td></tr>";
                content += "<!-- Zone -->";
                if (Img.Attributes.Contains("spectra_zone"))
                    TempName = Img.GetAttributeValue<EntityReference>("spectra_zone").Name;
                else
                    TempName = null;
                content += @"<tr style='background: #cccccc!important;'>
          <td colspan='2'>
            <p class='text-left'>Zone</p>
          </td>
          <td colspan='3'>
            <p class='text-left'>" + TempName + "</p></td></tr>";
                content += @"<!-- empty row  -->
        <tr>
          <td colspan='5'>&nbsp;</td>
        </tr>";
                content += @"<!-- Approval header -->
        <tr style='background: #cccccc!important;'>
          <td>
            <p class='text-center'><b>Material Name</b></p>
          </td>
          <td>
            <p class='text-center'><b>Item Code</b></p>
          </td>
          <td>
            <p><b>Expected Qty</b></p>
          </td>
          <td>
            <p><b>Consumed Qty</b></p>
          </td>
          <td>
            <p><b>Deviation</b></p>
          </td>
        </tr>";
                #region Checking Item consumptions: New Code created
                TempName = null;
                string Itemtable = "";
                fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='alletech_itemconsumption'>
                                    <attribute name='alletech_quantity' />
                                    <attribute name='alletech_actualquantityused' />
                                    <attribute name='alletech_subitem' />
                                    <attribute name='spectra_itemtype' />
                                    <order attribute='alletech_subitem' descending='false' />
                                    <filter type='and'>";
                if (Type == "WCR")
                {
                    fetch += "<condition attribute='alletech_wcr' operator='eq' uiname='' uitype='alletech_wcr' value='" + Img.Id.ToString() + @"' />";
                }
                else
                {
                    fetch += "<condition attribute='spectra_installationreport' operator='eq' uiname='' uitype='alletech_installationform' value='" + Img.Id.ToString() + @"' />";
                }

                fetch += @"       <condition attribute='spectra_itemtype' operator='eq' value='111260001' />
                                    </filter>
                                    <link-entity name='alletech_subitem' from='alletech_subitemid' to='alletech_subitem' visible='false' link-type='outer' alias='sub'>
                                      <attribute name='alletech_itemcodeinnav' />
                                    </link-entity>
                                  </entity>
                                </fetch>";

                TempColl = service.RetrieveMultiple(new FetchExpression(fetch));

                if (TempColl.Entities.Count > 0)
                {
                    string _duplicate_Item_Name = string.Empty;
                    int count = TempColl.Entities.Count;
                    for (int i = 0; i < count; i++)
                    {
                        try
                        {
                            #region New Code
                            if (TempColl.Entities[i].GetAttributeValue<OptionSetValue>("spectra_itemtype").Value == 111260001)
                            {
                                int eQty = 0, additional = 0, consumedQty = 0, deviation = 0;
                                string ItemCode = null;
                                string item = TempColl.Entities[i].GetAttributeValue<EntityReference>("alletech_subitem").Name;

                                if (_duplicate_Item_Name != item)
                                {
                                    if (TempColl.Entities[i].Attributes.Contains("sub.alletech_itemcodeinnav"))
                                        ItemCode = (string)TempColl.Entities[i].GetAttributeValue<AliasedValue>("sub.alletech_itemcodeinnav").Value;

                                    string ItemCodeCheck = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                      <entity name='alletech_itemconsumption'>
                                                        <attribute name='createdon' />
                                                        <attribute name='alletech_actualquantityused' />
                                                        <attribute name='spectra_itemtype' />
                                                        <attribute name='alletech_subitem' />
                                                        <attribute name='alletech_itemconsumptionid' />
                                                        <order attribute='spectra_itemtype' descending='false' />
                                                        <filter type='and'>";
                                    if (Type == "WCR")
                                    {
                                        ItemCodeCheck += "<condition attribute='alletech_wcr' operator='eq' uiname='' uitype='alletech_wcr' value='" + Img.Id.ToString() + @"' />";
                                    }
                                    else
                                    {
                                        ItemCodeCheck += "<condition attribute='spectra_installationreport' operator='eq' uiname='' uitype='alletech_installationform' value='" + Img.Id.ToString() + @"' />";
                                    }

                                    ItemCodeCheck += @" </filter>
                                                        <link-entity name='alletech_subitem' from='alletech_subitemid' to='alletech_subitem' alias='ac'>
                                                          <filter type='and'>
                                                            <condition attribute='alletech_itemcodeinnav' operator='eq' value='" + ItemCode + @"' />
                                                          </filter>
                                                        </link-entity>
                                                      </entity>
                                                    </fetch>";
                                    EntityCollection ItemCodeColl = service.RetrieveMultiple(new FetchExpression(ItemCodeCheck));
                                    if (ItemCodeColl.Entities.Count > 0)
                                    {
                                        foreach (Entity _default in ItemCodeColl.Entities)
                                        {
                                            if (_default.Attributes.Contains("spectra_itemtype"))
                                            {
                                                if (_default.GetAttributeValue<OptionSetValue>("spectra_itemtype").Value == 111260000)
                                                {
                                                    eQty += _default.GetAttributeValue<int>("alletech_actualquantityused");
                                                    //consumedQty += eQty;
                                                }
                                            }
                                        }
                                        foreach (Entity _additional in ItemCodeColl.Entities)
                                        {
                                            if (_additional.Attributes.Contains("spectra_itemtype"))
                                            {
                                                if (_additional.GetAttributeValue<OptionSetValue>("spectra_itemtype").Value == 111260001)
                                                {
                                                    additional += _additional.GetAttributeValue<int>("alletech_actualquantityused");
                                                    //consumedQty += eQty;
                                                }
                                            }
                                        }

                                    }
                                    _duplicate_Item_Name = TempColl.Entities[i].GetAttributeValue<EntityReference>("alletech_subitem").Name;
                                    consumedQty = eQty + additional;
                                    deviation = consumedQty - eQty;

                                    if (i % 2 == 0)
                                    {
                                        Itemtable += @"<tr ><td><p class='text-left'>" + item + "</p></td><td><p class='text-right'>" + ItemCode + "</p></td><td><p class='text-right'>" + eQty + "</p></td><td><p class='text-right'>" + consumedQty + "</p></td><td><p class='text-right'>" + deviation + "</p></td></tr>";
                                    }
                                    else
                                    {
                                        Itemtable += @"<tr style='background: #cccccc!important;'><td><p class='text-left'>" + item + "</p></td><td><p class='text-right'>" + ItemCode + "</p></td><td><p class='text-right'>" + eQty + "</p></td><td><p class='text-right'>" + consumedQty + "</p></td><td><p class='text-right'>" + deviation + "</p></td></tr>";

                                    }
                                    // Itemtable += @"<tr><td><p class='text-left'>" + item + "</p></td><td><p class='text-left'>" + ItemCode + "</p></td><td><p class='text-center'>" + eQty + "</p></td><td><p class='text-center'>" + consumedQty + "</p></td><td><p class='text-center'>" + deviation + "</p></td></tr>";
                                }
                            }
                            #endregion
                        }
                        catch (Exception ex)
                        {
                            throw new InvalidPluginExecutionException("Error in Item Consumption construct: " + ex.Message);
                        }
                    }
                }

                #endregion
                #region Checking Installation Item: New Code on 07-June-2021 by VLabs
                if (Type == "IR")
                {
                    TempName = null;
                    fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='alletech_installationitem'>
                                <attribute name='createdon' />
                                <attribute name='spectra_itemtype' />
                                <attribute name='alletech_installationitemid' />
                                <attribute name='alletech_subitem' />
                                <attribute name='spectra_quantityir' />
                                <order attribute='alletech_subitem' descending='false' />
                                <filter type='and'>
                                  <condition attribute='spectra_itemtype' operator='eq' value='111260001' />
                                  <condition attribute='alletech_installationform' operator='eq' value='" + Img.Id.ToString() + @"' />
                                </filter>
                                <link-entity name='alletech_subitem' from='alletech_subitemid' to='alletech_subitem' visible='false' link-type='outer' alias='sub'>
                                  <attribute name='alletech_itemcodeinnav' />
                                </link-entity>
                              </entity>
                            </fetch>";
                    TempColl = service.RetrieveMultiple(new FetchExpression(fetch));
                    if (TempColl.Entities.Count > 0)
                    {
                        string _duplicate_InstallationItem_Name = string.Empty;
                        int count = TempColl.Entities.Count;
                        for (int i = 0; i < count; i++)
                        {
                            try
                            {
                                //if additional
                                if (TempColl.Entities[i].GetAttributeValue<OptionSetValue>("spectra_itemtype").Value == 111260001)
                                {
                                    int eQty = 0, additional = 0, consumedQty = 0, deviation = 0;
                                    string ItemCode = null;
                                    string item = TempColl.Entities[i].GetAttributeValue<EntityReference>("alletech_subitem").Name;
                                    if (_duplicate_InstallationItem_Name != item)
                                    {
                                        if (TempColl.Entities[i].Attributes.Contains("sub.alletech_itemcodeinnav"))
                                            ItemCode = (string)TempColl.Entities[i].GetAttributeValue<AliasedValue>("sub.alletech_itemcodeinnav").Value;
                                        string _installationFetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                                      <entity name='alletech_installationitem'>
                                                                        <attribute name='createdon' />
                                                                        <attribute name='spectra_itemtype' />
                                                                        <attribute name='alletech_installationitemid' />
                                                                        <attribute name='alletech_subitem' />
                                                                        <attribute name='spectra_quantityir' />
                                                                        <order attribute='spectra_itemtype' descending='false' />
                                                                        <filter type='and'>
                                                                          <condition attribute='alletech_installationform' operator='eq' value='" + Img.Id.ToString() + @"' />
                                                                        </filter>
                                                                        <link-entity name='alletech_subitem' from='alletech_subitemid' to='alletech_subitem' alias='sub'>
                                                                          <attribute name='alletech_itemcodeinnav' />
                                                                          <filter type='and'>
                                                                            <condition attribute='alletech_itemcodeinnav' operator='eq' value='" + ItemCode + @"' />
                                                                          </filter>
                                                                        </link-entity>
                                                                      </entity>
                                                                    </fetch>";
                                        EntityCollection _installationItemCodeColl = service.RetrieveMultiple(new FetchExpression(_installationFetch));
                                        if (_installationItemCodeColl.Entities.Count > 0)
                                        {
                                            foreach (Entity _default in _installationItemCodeColl.Entities)
                                            {
                                                if (_default.Attributes.Contains("spectra_itemtype"))
                                                {
                                                    if (_default.GetAttributeValue<OptionSetValue>("spectra_itemtype").Value == 111260000)
                                                    {
                                                        eQty += _default.GetAttributeValue<int>("spectra_quantityir");
                                                    }
                                                }
                                            }
                                            foreach (Entity _additional in _installationItemCodeColl.Entities)
                                            {
                                                if (_additional.Attributes.Contains("spectra_itemtype"))
                                                {
                                                    if (_additional.GetAttributeValue<OptionSetValue>("spectra_itemtype").Value == 111260001)
                                                    {
                                                        additional += _additional.GetAttributeValue<int>("spectra_quantityir");
                                                    }
                                                }

                                                //if (_itemCode.Attributes.Contains("spectra_itemtype"))
                                                //{
                                                //    if (_itemCode.GetAttributeValue<OptionSetValue>("spectra_itemtype").Value == 111260000)
                                                //    {
                                                //        eQty += _itemCode.GetAttributeValue<int>("spectra_quantityir");
                                                //        consumedQty += eQty;
                                                //    }
                                                //    else
                                                //    {
                                                //        eQty = 0;
                                                //        consumedQty += _itemCode.GetAttributeValue<int>("spectra_quantityir");
                                                //    }
                                                //}
                                            }
                                            _duplicate_InstallationItem_Name = TempColl.Entities[i].GetAttributeValue<EntityReference>("alletech_subitem").Name;
                                            consumedQty = eQty + additional;
                                            deviation = consumedQty - eQty;
                                            if (i % 2 == 0)
                                            {
                                                Itemtable += @"<tr ><td><p class='text-left'>" + item + "</p></td><td><p class='text-right'>" + ItemCode + "</p></td><td><p class='text-right'>" + eQty + "</p></td><td><p class='text-right'>" + consumedQty + "</p></td><td><p class='text-right'>" + deviation + "</p></td></tr>";
                                            }
                                            else
                                            {
                                                Itemtable += @"<tr style='background: #cccccc!important;'><td><p class='text-left'>" + item + "</p></td><td><p class='text-right'>" + ItemCode + "</p></td><td><p class='text-right'>" + eQty + "</p></td><td><p class='text-right'>" + consumedQty + "</p></td><td><p class='text-right'>" + deviation + "</p></td></tr>";

                                            }
                                            // Itemtable += @"<tr><td><p class='text-left'>" + item + "</p></td><td><p class='text-left'>" + ItemCode + "</p></td><td><p class='text-center'>" + eQty + "</p></td><td><p class='text-center'>" + consumedQty + "</p></td><td><p class='text-center'>" + deviation + "</p></td></tr>";

                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new InvalidPluginExecutionException("Error in Installation Item Quantity adding in DOA: " + ex.Message);
                            }
                        }
                    }
                }
                #endregion
                content += Itemtable;//Items
                content += @"<!-- empty row after approval components list -->
        <tr>
          <!-- <td></td> -->
          <td colspan='5'>&nbsp;</td>
        </tr>
        <tr class='bg-grey'  style='background: #cccccc!important;'>
          <td colspan='5'>
            <p class='text-left fw-bold'>Remarks</p>
          </td>
        </tr>";
                if (Img.Attributes.Contains("alletech_remark"))
                    TempName = Img.GetAttributeValue<string>("alletech_remark");
                else
                    TempName = null;
                content += @"<tr>
          <td colspan='5'>
            <p>" + TempName + "</p></td></tr>";
                content += @"<!-- Note: -->
        <tr class='bg-grey'  style='background: #cccccc!important;'>
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

        public string GetCustomersegment(int value)
        {
            string text = null;

            switch (value)
            {
                case 111260000: text = "SMB"; break;
                case 111260001: text = "Media"; break;
                case 111260002: text = "LA"; break;
                case 111260003: text = "SP"; break;
            }
            return text;
        }

        public string getBuildingStatusValue(int value)
        {
            string text = null;

            switch (value)
            {
                case 1: text = "Non-RFS"; break;
                case 2: text = "B-RFS"; break;
                case 3: text = "P-RFS"; break;
                case 4: text = "C-RFS"; break;
                case 5: text = "A-RFS Type1"; break;
                case 6: text = "A-RFS Type2"; break;
                case 7: text = "Pb-RFS"; break;
                case 8: text = "L2P-RFS"; break;
                case 9: text = "L2B-RFS"; break;
            }
            return text;
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

        public Entity GetResultByAttribute(IOrganizationService _service, string entityName, string attrName, string attrValue, string[] column)
        {
            try
            {
                Entity result = null;
                QueryExpression query = new QueryExpression(entityName);
                query.NoLock = true;
                query.ColumnSet.AddColumns(column);
                query.Criteria.AddCondition(attrName, ConditionOperator.Equal, attrValue);
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

        public EntityCollection getApprovalConfig(IOrganizationService service, string appConfigNameType)//, decimal percentAge)
        {
            QueryExpression query = new QueryExpression("spectra_approvalconfig");
            query.ColumnSet = new ColumnSet("spectra_approver", "spectra_name", "spectra_orderby", "spectra_minpercentage", "spectra_maxpercentage", "spectra_quantity");//spectra_percentage
            query.Criteria.AddCondition(new ConditionExpression("spectra_name", ConditionOperator.Equal, appConfigNameType.ToUpper()));
            query.Criteria.AddCondition(new ConditionExpression("spectra_productsegment", ConditionOperator.Null));
            query.Orders.Add(new OrderExpression("spectra_minpercentage", OrderType.Ascending));
            query.Orders.Add(new OrderExpression("spectra_orderby", OrderType.Ascending));
            return service.RetrieveMultiple(query);
        }
    }
}
