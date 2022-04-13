using System;
using System.Collections.Generic;
using log4net;
using Newtonsoft.Json;
using System.Runtime.InteropServices;
using SAPbouiCOM;
using System.Resources;
using System.Reflection;

namespace T1.B1.SelfWithholdingTax
{
    public class Operations
    {
        private static Operations objSelfWithHoldingTax;
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static bool runResizelogic = true;
        private static List<string> WHPurchaseDocuments = new List<string>();
        private static List<string> WHSalesDocuments = new List<string>();
        private Operations()
        {
            WHPurchaseDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._SelfWithHoldingTax.WTPurchaseObjects);
            WHSalesDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._SelfWithHoldingTax.WTSalesObjects);
        }
        public static void formDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool blBubbleEvent)
        {
            WHPurchaseDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._SelfWithHoldingTax.WTPurchaseObjects);
            WHSalesDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._SelfWithHoldingTax.WTSalesObjects);

            if (objSelfWithHoldingTax == null) objSelfWithHoldingTax = new Operations();

            try
            {
                if (!BusinessObjectInfo.BeforeAction && BusinessObjectInfo.ActionSuccess)
                {

                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            if (WHPurchaseDocuments.Contains(BusinessObjectInfo.FormTypeEx) || WHSalesDocuments.Contains(BusinessObjectInfo.FormTypeEx))
                            {
                                SelfWithholdingTax.addSelfWithHoldingTax(BusinessObjectInfo);
                            }
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            if (WHPurchaseDocuments.Contains(BusinessObjectInfo.FormTypeEx) || WHSalesDocuments.Contains(BusinessObjectInfo.FormTypeEx))
                            {
                                SelfWithholdingTax.getSWTaxInfoForDocument(BusinessObjectInfo);
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
            EventInfoClass eventInfo = null;

            try
            {
                if (!pVal.BeforeAction)
                {
                    switch (pVal.MenuUID)
                    {
                        case "HCO_MSW0002":
                            SelfWithholdingTax.loadSWTaxConfigForm();
                            break;
                        case "HCO_MSW0003":
                            SelfWithholdingTax.loadMissingSWTaxForm();
                            break;
                        case "HCO_MSW0004":
                            SelfWithholdingTax.loadCancelSWTaxForm();
                            break;
                        case "HCO_MWTRU":
                            eventInfo = CacheManager.CacheManager.Instance.getFromCache(Settings._Main.lastRightClickEventInfo);
                            SelfWithholdingTax.relatedPartiedMatrixOperationUDO(eventInfo, "Add");
                            break;
                        case "HCO_MWTDRU":
                            eventInfo = CacheManager.CacheManager.Instance.getFromCache(Settings._Main.lastRightClickEventInfo);
                            SelfWithholdingTax.relatedPartiedMatrixOperationUDO(eventInfo, "Delete");
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
            string[] showInFolderList;
            bool blInList = false;
            try
            {
                if (!pVal.BeforeAction)
                {
                    if (pVal.ActionSuccess)
                    {
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_CHOOSE_FROM_LIST:
                                if (pVal.FormTypeEx == "HCO_FSW0100")
                                {
                                    if (pVal.ItemUID == "13_U_E") SelfWithholdingTax.clearfilterAccounts(pVal, "CFL_DB");
                                    if (pVal.ItemUID == "14_U_E") SelfWithholdingTax.clearfilterAccounts(pVal, "CFL_CR");
                                    if (pVal.ItemUID == "0_U_G")
                                    {
                                        SelfWithholdingTax.setBPNameColumn(pVal);
                                        SelfWithholdingTax.clearfilterBPs(pVal);
                                    }
                                }
                                break;
                            case BoEventTypes.et_FORM_RESIZE:
                                if (runResizelogic)
                                {
                                    showInFolderList = Settings._SelfWithHoldingTax.showFolderInDocumentsList.Split(',');
                                    for (int i = 0; i < showInFolderList.Length; i++)
                                    {
                                        if (showInFolderList[i] == pVal.FormTypeEx)
                                        {
                                            blInList = true;
                                            break;
                                        }
                                    }
                                    if (blInList)
                                    {
                                        SelfWithholdingTax.HCOSelfWithHoldingFolderAdd(pVal.FormUID);
                                        blInList = false;
                                    }

                                }
                                runResizelogic = true;
                                break;
                            case BoEventTypes.et_COMBO_SELECT:
                                break;
                            case BoEventTypes.et_ITEM_PRESSED:
                                if (pVal.FormTypeEx == Settings._SelfWithHoldingTax.CancelFormUID)
                                {
                                    if (pVal.ItemUID == "1") SelfWithholdingTax.getPostedSWTaxDocuments(FormUID, pVal);
                                    if (pVal.ItemUID == "txtSWTCode") SelfWithholdingTax.setSelectedCode(pVal);
                                    if (pVal.ItemUID == "btnCalc") SelfWithholdingTax.cancelPostedTaxDocuments(FormUID, pVal);
                                    if (pVal.FormTypeEx == "HCO_FSW0100")
                                    {
                                        if (pVal.ItemUID == "btnAddAll")
                                        {
                                            SelfWithholdingTax.addAllPBS(pVal);
                                        }
                                        if (pVal.ItemUID == "btnClear")
                                        {
                                            SelfWithholdingTax.clearAllPBS(pVal);
                                        }
                                    }
                                    if(pVal.FormTypeEx == Settings._SelfWithHoldingTax.MissingSWTFormUID)
                                    {
                                        if (pVal.ItemUID == "btnCalc") SelfWithholdingTax.addMisingSWTDocuments(FormUID, pVal); 
                                    }
                                }
                                if (pVal.ItemUID == Settings._SelfWithHoldingTax.SelfWithHoldingFolderId)
                                {
                                    showInFolderList = Settings._SelfWithHoldingTax.showFolderInDocumentsList.Split(',');
                                    for (int i = 0; i < showInFolderList.Length; i++)
                                    {
                                        if (showInFolderList[i] == pVal.FormTypeEx)
                                        {
                                            blInList = true;
                                            break;
                                        }
                                    }
                                    if (blInList)
                                    {
                                        MainObject.Instance.B1Application.Forms.Item(pVal.FormUID).PaneLevel = Settings._SelfWithHoldingTax.SelfWithHoldingFolderPane;
                                    }
                                }
                                if(pVal.FormTypeEx == Settings._SelfWithHoldingTax.MissingSWTFormUID)
                                {
                                    if(pVal.ItemUID == "btnGet") SelfWithholdingTax.getMissingSWTaxDocuments(FormUID, pVal);

                                }
                                break;
                            case BoEventTypes.et_DOUBLE_CLICK:
                                if (pVal.FormTypeEx == Settings._SelfWithHoldingTax.CancelFormUID)
                                {
                                    if (pVal.ItemUID == "grdSWT") T1.B1.Base.UIOperations.Operations.toggleSelectCheckBox(pVal, "dtSelfWT", "1");
                                }
                                if (pVal.FormTypeEx == Settings._SelfWithHoldingTax.MissingSWTFormUID)
                                {
                                    if (pVal.ItemUID == "grdSWT") T1.B1.Base.UIOperations.Operations.toggleSelectCheckBox(pVal, "dtSelfWT", "1");
                                }
                                break;
                            case BoEventTypes.et_FORM_LOAD:
                                showInFolderList = Settings._SelfWithHoldingTax.showFolderInDocumentsList.Split(',');
                                for (int i = 0; i < showInFolderList.Length; i++)
                                {
                                    if (showInFolderList[i] == pVal.FormTypeEx)
                                    {
                                        blInList = true;
                                        break;
                                    }
                                }
                                if (blInList)
                                {

                                    SelfWithholdingTax.HCOSelfWithHoldingFolderAdd(pVal.FormUID);
                                    runResizelogic = false;
                                }
                                break;
                        }
                    }
                }
                else
                {
                    switch (pVal.EventType)
                    {
                        case BoEventTypes.et_ITEM_PRESSED:
                            if (pVal.FormTypeEx == "HCO_FSW0100")
                            {
                                if (pVal.ItemUID == "1")
                                {
                                    Form oform = MainObject.Instance.B1Application.Forms.ActiveForm;
                                    Matrix oMatriz = (Matrix)oform.Items.Item("0_U_G").Specific;
                                    oMatriz.FlushToDataSource();
                                    if(oform.Mode != BoFormMode.fm_FIND_MODE)
                                    {
                                        var lines = oform.DataSources.DBDataSources.Item(1).Size;
                                        for (int i = 0; i < lines; i++)
                                        {
                                            if (oform.DataSources.DBDataSources.Item(1).GetValue("U_CardCode", i).Equals(string.Empty)) oform.DataSources.DBDataSources.Item(1).RemoveRecord(i);
                                        }
                                    }

                                }
                            }
                                break;
                        case BoEventTypes.et_CHOOSE_FROM_LIST:
                            if (pVal.FormTypeEx == "HCO_FSW0100")
                            {
                                if (pVal.ItemUID == "13_U_E")
                                {
                                    SelfWithholdingTax.filterAccounts(pVal, "CFL_DB");
                                }
                                if (pVal.ItemUID == "14_U_E")
                                {
                                    SelfWithholdingTax.filterAccounts(pVal, "CFL_CR");
                                }
                                if (pVal.ItemUID == "0_U_G")
                                {
                                    SelfWithholdingTax.filterBPs(pVal);
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
        }

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
                        case "HCO_FSW0100":
                            if (eventInfo.ItemUID == "0_U_G")
                            {
                                SelfWithholdingTax.addInsertRowRelationMenuUDO(objForm, eventInfo);
                                SelfWithholdingTax.addDeleteRowRelationMenuUDO(objForm, eventInfo);
                            }
                            break;
                        case "133":
                            CacheManager.CacheManager.Instance.addToCache("LastActiveForm", objForm.UniqueID, CacheManager.CacheManager.objCachePriority.Default);
                            break;
                    }
                }
                else
                {
                    switch (objForm.TypeEx)
                    {
                        case "HCO_FSW0100":
                            if (eventInfo.ItemUID == "0_U_G")
                            {
                                SelfWithholdingTax.removeDeleteRowRelationMenuUDO();
                                SelfWithholdingTax.removeInsertRowRelationMenuUDO();
                            }
                            break;
                        case "133":
                            CacheManager.CacheManager.Instance.removeFromCache("LastActiveForm");
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
