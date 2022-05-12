using System;
using System.Collections.Generic;
using log4net;
using Newtonsoft.Json;
using System.Runtime.InteropServices;
using SAPbouiCOM;
using System.Resources;
using System.Reflection;
using System.Xml;

namespace T1.B1.WithholdingTax
{
    public enum FORM_MODE { SEARCH, NEW, OK}
    public class Operations
    {
        private static Operations objWithHoldingTax;
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static bool runResizelogic = true;
        private static List<string> WHPurchaseDocuments = new List<string>();
        private static List<string> WHSalesDocuments = new List<string>();
        public static bool CloseFormBP = false;
        private Operations()
        {
            try
            {
                WHPurchaseDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._WithHoldingTax.WTPurchaseObjects);
                WHSalesDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._WithHoldingTax.WTSalesObjects);
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }

        }
        public static void formDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool blBubbleEvent)
        {
            if (objWithHoldingTax == null) objWithHoldingTax = new Operations();
            XmlDocument oXml = new XmlDocument();
           
            try
            {
                if (!BusinessObjectInfo.BeforeAction)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            if (BusinessObjectInfo.ActionSuccess && (WHPurchaseDocuments.Contains(BusinessObjectInfo.FormTypeEx) || WHSalesDocuments.Contains(BusinessObjectInfo.FormTypeEx)))
                            {
                                AddDocumentInfoArgs objArgs = new AddDocumentInfoArgs();
                                oXml.LoadXml(BusinessObjectInfo.ObjectKey);
                                XmlNode Xn = oXml.LastChild;                                
                                objArgs.ObjectKey = Xn["DocEntry"].InnerText;
                                objArgs.ObjectType = BusinessObjectInfo.Type;
                                objArgs.FormtTypeEx = BusinessObjectInfo.FormTypeEx;
                                objArgs.FormUID = BusinessObjectInfo.FormUID;

                                WithholdingTax.addDocumentInfo(objArgs);
                            }
                            break;
                        case BoEventTypes.et_FORM_UNLOAD:
                            if (BusinessObjectInfo.ActionSuccess && (WHPurchaseDocuments.Contains(BusinessObjectInfo.FormTypeEx) || WHSalesDocuments.Contains(BusinessObjectInfo.FormTypeEx)))
                            {
                                CacheManager.CacheManager.Instance.removeFromCache("Disable_" + BusinessObjectInfo.FormUID);
                                CacheManager.CacheManager.Instance.removeFromCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + BusinessObjectInfo.FormUID);
                                CacheManager.CacheManager.Instance.removeFromCache(Settings._WithHoldingTax.WTFOrmInfoCachePrefix + BusinessObjectInfo.FormUID);
                                CacheManager.CacheManager.Instance.removeFromCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + BusinessObjectInfo.FormUID);
                                CacheManager.CacheManager.Instance.removeFromCache("WTLogicDone_" + BusinessObjectInfo.FormUID);
                                CacheManager.CacheManager.Instance.removeFromCache("WTCFLExecuted");
                                CacheManager.CacheManager.Instance.removeFromCache("WTLastActiveForm");
                            }
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            if (BusinessObjectInfo.FormTypeEx == "HCO_FWT1100")
                            {
                                SAPbouiCOM.Form objForm = MainObject.Instance.B1Application.Forms.Item(BusinessObjectInfo.FormUID);
                                LinkedButton oLink = (SAPbouiCOM.LinkedButton)objForm.Items.Item("Item_13").Specific;
                                oLink.LinkedObject = (SAPbouiCOM.BoLinkedObject)(Int32.Parse(objForm.DataSources.DBDataSources.Item(0).GetValue("U_DocType", 0)));
                                oLink.Item.Visible = true;
                                oLink = (SAPbouiCOM.LinkedButton)objForm.Items.Item("Item_11").Specific;
                                oLink.Item.Visible = true;

                            }
                            break;
                    }
                }
                else
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            if (WHPurchaseDocuments.Contains(BusinessObjectInfo.FormTypeEx) || WHSalesDocuments.Contains(BusinessObjectInfo.FormTypeEx))
                            {
                                Form oForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                                if (!WithholdingTax.HasRelParty(oForm.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0)))
                                {
                                    MainObject.Instance.B1Application.MessageBox("El " + (WHPurchaseDocuments.Contains(BusinessObjectInfo.FormTypeEx) ? "proveedor " : "cliente ") + "no tiene tercero relacionado");
                                    blBubbleEvent = false;
                                }
                            }
                            break;
                    }

                }
            }
            catch (COMException COMException)
            {
                _Logger.Error("", COMException);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }
        public static void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (!pVal.BeforeAction)
                {
                    switch (pVal.MenuUID)
                    {
                        case "HCO_MWT0002":
                            WithholdingTax.LoadWithHoldingForm(Settings.EWithHoldingTax.MUNICIPALITY);
                            break;
                        case "HCO_MWT0003":
                            WithholdingTax.LoadWithHoldingForm(Settings.EWithHoldingTax.WITHHOLDING_OPERATION);
                            SAPbouiCOM.Form objForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                            LinkedButton oLink = (SAPbouiCOM.LinkedButton)objForm.Items.Item("Item_11").Specific;
                            oLink.LinkedObjectType = "HCO_FRP1100";
                            oLink.LinkedFormXmlPath = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Forms\\HCO_Terceros_Relacionados.srf";
                            oLink.Item.Visible = false;
                            break;
                        case "HCO_MWTARU":
                            var eventInfoMrparu = CacheManager.CacheManager.Instance.getFromCache(Settings._Main.lastRightClickEventInfo);
                            WithholdingTax.MatrixOperationUDO("Add", "0_U_G");
                            WithholdingTax.rowNumber("0_U_G");
                            break;
                        case "HCO_MWT0004":
                            WithholdingTax.LoadWithHoldingForm(Settings.EWithHoldingTax.MISSING_OPERATIONS);
                            WithholdingTax.InitMissingOperationsForm();
                            break;
                        case "1281":
                            var form = MainObject.Instance.B1Application.Forms.ActiveForm;
                            if (form.TypeEx == "HCO_FWT1100")
                            {
                                WithholdingTax.SetFormState(form, FORM_MODE.SEARCH);
                            }
                            break;
                    }
                }
                else
                {
                    string strLastActiveForm = string.Empty;
                    switch (pVal.MenuUID)
                    {
                        case "5897":
                            strLastActiveForm = MainObject.Instance.B1Application.Forms.ActiveForm.UniqueID;
                            CacheManager.CacheManager.Instance.addToCache("WTLastActiveForm", strLastActiveForm, CacheManager.CacheManager.objCachePriority.Default);
                            break;

                        case "6005":
                            strLastActiveForm = MainObject.Instance.B1Application.Forms.ActiveForm.UniqueID;
                            CacheManager.CacheManager.Instance.addToCache("LastActiveForm", strLastActiveForm, CacheManager.CacheManager.objCachePriority.Default);
                            break;
                    }
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }
        public static void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {

            WHPurchaseDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._WithHoldingTax.WTPurchaseObjects);
            WHSalesDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._WithHoldingTax.WTSalesObjects);
            
            try
            {

                if (!pVal.BeforeAction)
                {
                    switch (pVal.EventType)
                    {
                        case BoEventTypes.et_CHOOSE_FROM_LIST:
                            if (WHPurchaseDocuments.Contains(pVal.FormTypeEx) || WHSalesDocuments.Contains(pVal.FormTypeEx))
                            {
                                if (pVal.ItemUID == "4" || pVal.ItemUID == "54")
                                {
                                    if (WithholdingTax.formModeAdd(pVal))
                                    {
                                        WithholdingTax.getSelectedBPInformation(pVal, true);
                                        CacheManager.CacheManager.Instance.addToCache("WTCFLExecuted", true, CacheManager.CacheManager.objCachePriority.Default);
                                    }
                                    CacheManager.CacheManager.Instance.addToCache("WTLastActiveForm", MainObject.Instance.B1Application.Forms.ActiveForm.UniqueID, CacheManager.CacheManager.objCachePriority.Default);
                                }
                            }
                            if(pVal.FormTypeEx == "HCO_FWT0100")
                            {
                                WithholdingTax.SetChooseFromListMunMatrix(pVal);
                            }
                            break;

                        case BoEventTypes.et_LOST_FOCUS:
                            
                        if (WHPurchaseDocuments.Contains(pVal.FormType.ToString()) || WHSalesDocuments.Contains(pVal.FormType.ToString()))
                            {
                                if (WithholdingTax.formModeAdd(pVal))
                                {

                                    bool WTExec = CacheManager.CacheManager.Instance.getFromCache("WTCFLExecuted") == null ? false : true;
                                    bool LogicDone = CacheManager.CacheManager.Instance.getFromCache("WTLogicDone_" + pVal.FormUID) == null ? false : true;
                                    if (pVal.ItemUID == "4" || pVal.ItemUID == "54" || pVal.ItemUID == "40")
                                    {
                                        LogicDone = false;
                                    }

                                    if (!LogicDone)
                                    {
                                        if (!WTExec) WithholdingTax.getSelectedBPInformation(pVal, false);
                                        else CacheManager.CacheManager.Instance.removeFromCache("WTCFLExecuted");
                                        Form oForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                                        if (!oForm.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0).Equals("") && !oForm.DataSources.DBDataSources.Item(0).GetValue("DocDate", 0).Equals(""))
                                            WithholdingTax.activateWTMenu(pVal.FormUID, true);
                                    }
                                }
                            }
                            break;

                        case BoEventTypes.et_COMBO_SELECT:
                            if (WHPurchaseDocuments.Contains(pVal.FormTypeEx) || WHSalesDocuments.Contains(pVal.FormTypeEx))
                            {
                                if (pVal.ItemUID.Equals("226"))
                                {
                                    if (WithholdingTax.formModeAdd(pVal))
                                    {
                                        WithholdingTax.getSelectedBPInformation(pVal, false);
                                        Form oForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                                        if (!oForm.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0).Equals(""))
                                            WithholdingTax.activateWTMenu(pVal.FormUID, true);
                                    }
                                }
                            }
                            break;

                        case BoEventTypes.et_ITEM_PRESSED:
                            if (WHPurchaseDocuments.Contains(pVal.FormTypeEx) || WHSalesDocuments.Contains(pVal.FormTypeEx))
                            {
                                if (pVal.ItemUID.Equals("173"))
                                {
                                    if (WithholdingTax.formModeAdd(pVal)) CacheManager.CacheManager.Instance.addToCache("WTLastActiveForm", MainObject.Instance.B1Application.Forms.ActiveForm.UniqueID, CacheManager.CacheManager.objCachePriority.Default);
                                }

                            }

                            break;
                        case BoEventTypes.et_FORM_LOAD:
                            if (pVal.FormTypeEx.Equals("60504"))
                            {
                                string strLastActiveForm = CacheManager.CacheManager.Instance.getFromCache("WTLastActiveForm") == null ? "" : CacheManager.CacheManager.Instance.getFromCache("WTLastActiveForm");
                                if (strLastActiveForm.Trim().Length > 0)
                                {
                                    bool blDisabled = CacheManager.CacheManager.Instance.getFromCache("Disable_" + strLastActiveForm) != null ? true : false;
                                    if (!blDisabled)
                                    {
                                        string strFormAutoActivate = CacheManager.CacheManager.Instance.getFromCache("WTAutoActivate") != null ? CacheManager.CacheManager.Instance.getFromCache("WTAutoActivate") : "";
                                        if (strFormAutoActivate.Trim() == strLastActiveForm.Trim())
                                        {
                                            T1.B1.Base.UIOperations.Operations.startProgressBar("Asignando retenciones automáticas...", 2);
                                            WithholdingTax.setBPWT(strFormAutoActivate, pVal);
                                        }
                                    }
                                }
                            }
                            if (pVal.FormTypeEx.Equals("frmDummy"))
                            {
                                if (CloseFormBP)
                                {
                                    var asdas = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                                    asdas.Left = -1000;
                                    asdas.Top = -2312;
                                    asdas.VisibleEx = false;
                                }
                            }
                            break;
                        case BoEventTypes.et_FORM_UNLOAD:
                            if (WHPurchaseDocuments.Contains(pVal.FormTypeEx) || WHSalesDocuments.Contains(pVal.FormTypeEx))
                            {
                                CacheManager.CacheManager.Instance.removeFromCache("Disable_" + pVal.FormUID);
                                CacheManager.CacheManager.Instance.removeFromCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + pVal.FormUID);
                                CacheManager.CacheManager.Instance.removeFromCache(Settings._WithHoldingTax.WTFOrmInfoCachePrefix + pVal.FormUID);
                                CacheManager.CacheManager.Instance.removeFromCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + pVal.FormUID);
                                CacheManager.CacheManager.Instance.removeFromCache("WTLogicDone_" + pVal.FormUID);
                                CacheManager.CacheManager.Instance.removeFromCache("WTCFLExecuted");
                                CacheManager.CacheManager.Instance.removeFromCache("WTLastActiveForm");
                            }
                            break;

                    }
                }
                else
                {
                    switch (pVal.EventType)
                    {
                        case BoEventTypes.et_ITEM_PRESSED:
                            if (WHPurchaseDocuments.Contains(pVal.FormTypeEx) || WHSalesDocuments.Contains(pVal.FormTypeEx))
                            {
                                if (pVal.ItemUID.Equals("1"))
                                {
                                    Form oForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                                    if (!WithholdingTax.HasRelParty(oForm.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0)))
                                    {
                                        MainObject.Instance.B1Application.MessageBox("El " + (WHPurchaseDocuments.Contains(pVal.FormTypeEx) ? "proveedor " : "cliente ") + "no tiene tercero relacionado");
                                        BubbleEvent = false;
                                    }
                                }
                            }

                            if (pVal.FormTypeEx.Equals("HCO_FWT1200"))
                            {
                                if (pVal.ItemUID == "btnAdd")
                                {
                                    Form oForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                                    if (oForm.PaneLevel == 1)
                                        WithholdingTax.createMissingOperations(pVal);
                                    else
                                        oForm.Close();
                                }
                                    
                            }

                                if (pVal.FormTypeEx.Equals("60504"))
                            {
                                if (pVal.ItemUID == "1")
                                {
                                    string blAutoActivate = CacheManager.CacheManager.Instance.getFromCache("WTAutoActivate") != null ? CacheManager.CacheManager.Instance.getFromCache("WTAutoActivate") : "";
                                    SAPbouiCOM.Form objForm = null;
                                    string strLastActiveForm = CacheManager.CacheManager.Instance.getFromCache("WTLastActiveForm") == null ? "" : CacheManager.CacheManager.Instance.getFromCache("WTLastActiveForm");
                                    if (strLastActiveForm.Trim().Length > 0)
                                    {
                                        bool isDisabled = CacheManager.CacheManager.Instance.getFromCache("Disable_" + strLastActiveForm) == null ? false : true;
                                        if (blAutoActivate.Trim().Length == 0)
                                        {
                                            objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                                            if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE && !isDisabled)
                                            {
                                                if (MainObject.Instance.B1Application.MessageBox("La modificación manual de las retenciones deshabilitará el cálculo automatico para este documento. Desea Continuar? ", 2, "Sí", "No", "") != 2)
                                                {
                                                    if (strLastActiveForm.Trim().Length > 0)
                                                    {
                                                        CacheManager.CacheManager.Instance.addToCache(string.Concat("Disable_", strLastActiveForm), true, CacheManager.CacheManager.objCachePriority.Default);
                                                    }
                                                    else
                                                    {
                                                        if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                                        {
                                                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                                        }
                                                        else BubbleEvent = false;
                                                    }

                                                    objForm.Close();
                                                }
                                                else
                                                {
                                                    BubbleEvent = false;
                                                    objForm.Items.Item("2").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                                }
                                            }
                                            CacheManager.CacheManager.Instance.removeFromCache("WTLastActiveForm");
                                            T1.B1.Base.UIOperations.Operations.stopProgressBar();

                                        }
                                    }
                                }
                            }


                            break;
                    }
                }

            }
            catch (COMException comEx)
            {
                _Logger.Error("", comEx);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                T1.B1.Base.UIOperations.Operations.stopProgressBar();
            }
        }

        //public static void RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        //{
        //    if (objWithHoldingTax == null)
        //    {
        //        objWithHoldingTax = new Operations();
        //    }

        //    SAPbouiCOM.Form objForm = null;
        //    BubbleEvent = true;
        //    try
        //    {
        //        objForm = MainObject.Instance.B1Application.Forms.Item(eventInfo.FormUID);

        //        #region UDO Form
        //        if (objForm.TypeEx == "HCO_T1SWT100UDO"
        //            && eventInfo.BeforeAction
        //            && eventInfo.ItemUID == "0_U_G"

        //            )
        //        {
        //            SelfWithholdingTax.addInsertRowRelationMenuUDO(objForm, eventInfo);
        //            SelfWithholdingTax.addDeleteRowRelationMenuUDO(objForm, eventInfo);



        //        }

        //        if (objForm.TypeEx == "HCO_T1SWT100UDO"
        //            && !eventInfo.BeforeAction
        //            && eventInfo.ItemUID == "0_U_G"

        //            )
        //        {
        //            SelfWithholdingTax.removeDeleteRowRelationMenuUDO();
        //            SelfWithholdingTax.removeInsertRowRelationMenuUDO();


        //        }
        //        #endregion
        //        #region Invoice
        //        if(objForm.TypeEx == "133"
        //            && eventInfo.BeforeAction
        //            )
        //        {
        //            string strLastActiveForm = MainObject.Instance.B1Application.Forms.ActiveForm.UniqueID;
        //            CacheManager.CacheManager.Instance.addToCache("LastActiveForm", strLastActiveForm, CacheManager.CacheManager.objCachePriority.Default);

        //        }

        //        if(objForm.TypeEx == "133"
        //            && !eventInfo.BeforeAction)
        //        {
        //            CacheManager.CacheManager.Instance.removeFromCache("LastActiveForm");

        //        }
        //        #endregion
        //    }
        //    catch (COMException comEx)
        //    {
        //        Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
        //        _Logger.Error("", comEx);

        //    }
        //    catch (Exception er)
        //    {
        //        _Logger.Error("", er);
        //    }
        //}

        public static void RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        {
            SAPbouiCOM.Form objForm = null;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(eventInfo.FormUID);
                if (eventInfo.BeforeAction)
                {
                    switch (objForm.TypeEx)
                    {
                        case "HCO_FWT0100":
                            if (eventInfo.ItemUID == "0_U_G")
                            {
                                WithholdingTax.addInsertRowRelationMenuUDO(objForm, eventInfo);
                                WithholdingTax.addDeleteRowRelationMenuUDO(objForm, eventInfo);
                            }

                            MainObject.Instance.B1Application.Menus.Item("1283").Enabled = false;
                            MainObject.Instance.B1Application.Menus.Item("1284").Enabled = false;
                            break;
                    }
                }
                else
                {
                    switch (objForm.TypeEx)
                    {
                        case "HCO_FWT0100":
                            if (eventInfo.ItemUID == "Item_1")
                            {
                                WithholdingTax.removeDeleteRowRelationMenuUDO();
                                WithholdingTax.removeInsertRowRelationMenuUDO();
                            }



                            break;
                    }
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }



    }
}
