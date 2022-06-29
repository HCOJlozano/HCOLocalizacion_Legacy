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
        private static List<string> SWTDocuments = new List<string>();
        private static Form objForm;

        private Operations()
        {
            SWTDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._SelfWithHoldingTax.SWTDocuments);
        }
        public static void formDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool blBubbleEvent)
        {
            SWTDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._SelfWithHoldingTax.SWTDocuments);

            if (objSelfWithHoldingTax == null) objSelfWithHoldingTax = new Operations();

            try
            {
                if (!BusinessObjectInfo.BeforeAction && BusinessObjectInfo.ActionSuccess)
                {

                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            if (SWTDocuments.Contains(BusinessObjectInfo.FormTypeEx))
                            {
                                SelfWithholdingTax.addSelfWithHoldingTax(BusinessObjectInfo);
                            }
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            if (SWTDocuments.Contains(BusinessObjectInfo.FormTypeEx))
                            {
                                SelfWithholdingTax.getSWTaxInfoForDocument(BusinessObjectInfo);
                            }
                            else if (BusinessObjectInfo.FormTypeEx == "HCO_FSW0100")
                            {
                                objForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                                SelfWithholdingTax.EnableItems(false, objForm);
                                SelfWithholdingTax.CheckGroups(objForm);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objForm);
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
                            SelfWithholdingTax.MatrixOperationUDO(eventInfo, "Add");
                            break;
                        case "HCO_MWTDRU":
                            eventInfo = CacheManager.CacheManager.Instance.getFromCache(Settings._Main.lastRightClickEventInfo);
                            SelfWithholdingTax.MatrixOperationUDO(eventInfo, "Delete");
                            break;
                        case "1282":
                            objForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                            SelfWithholdingTax.EnableItems(true, objForm);
                            SelfWithholdingTax.AddGroupsToMatrix(objForm);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objForm);
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
                                    SelfWithholdingTax.SelectFieldCFL(pVal);
                                break;
                            case BoEventTypes.et_FORM_RESIZE:
                                //if (runResizelogic)
                                //{
                                //    showInFolderList = Settings._SelfWithHoldingTax.showFolderInDocumentsList.Split(',');
                                //    for (int i = 0; i < showInFolderList.Length; i++)
                                //    {
                                //        if (showInFolderList[i] == pVal.FormTypeEx)
                                //        {
                                //            blInList = true;
                                //            break;
                                //        }
                                //    }
                                //    if (blInList)
                                //    {
                                //        SelfWithholdingTax.HCOSelfWithHoldingFolderAdd(pVal.FormUID);
                                //        blInList = false;
                                //    }

                                //}
                                //runResizelogic = true;
                                break;
                            case BoEventTypes.et_COMBO_SELECT:
                                if (pVal.FormTypeEx == "HCO_FSW0100")
                                {
                                    if (pVal.ItemUID == "Item_30")
                                    {
                                        objForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                                        var tipo = ((ComboBox)objForm.Items.Item(pVal.ItemUID).Specific).Value;
                                        switch (tipo)
                                        {
                                            case "S":
                                                ((EditText)objForm.Items.Item("Item_16").Specific).Value = "";
                                                objForm.Items.Item("Item_8").Click(BoCellClickType.ct_Regular);
                                                objForm.Items.Item("Item_16").Enabled = false;
                                                break;
                                            case "G":
                                                objForm.Items.Item("Item_16").Enabled = true;
                                                break;
                                        }
                                    }
                                }
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
                                    if (pVal.FormTypeEx == Settings._SelfWithHoldingTax.MissingSWTFormUID)
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
                                if (pVal.FormTypeEx == Settings._SelfWithHoldingTax.MissingSWTFormUID)
                                {
                                    if (pVal.ItemUID == "btnGet") SelfWithholdingTax.getMissingSWTaxDocuments(FormUID, pVal);

                                //}
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
                                //showInFolderList = Settings._SelfWithHoldingTax.showFolderInDocumentsList.Split(',');
                                //for (int i = 0; i < showInFolderList.Length; i++)
                                //{
                                //    if (showInFolderList[i] == pVal.FormTypeEx)
                                //    {
                                //        blInList = true;
                                //        break;
                                //    }
                                //}
                                //if (blInList)
                                //{

                                //    SelfWithholdingTax.HCOSelfWithHoldingFolderAdd(pVal.FormUID);
                                //    runResizelogic = false;
                                //}
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
                                    BubbleEvent = SelfWithholdingTax.ValidateFields(pVal);
                                }
                            }

                            break;

                        case BoEventTypes.et_MATRIX_LINK_PRESSED:

                            if (SWTDocuments.Contains(pVal.FormTypeEx))
                            {
                                if (pVal.ColUID == "Nro Reg. Retencion")
                                    SelfWithholdingTax.OpenSelfWitholdingRecord(pVal);
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
                            if (eventInfo.ItemUID == "Item_31" || eventInfo.ItemUID == "Item_12")
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
                            if (eventInfo.ItemUID == "Item_31" || eventInfo.ItemUID == "Item_12")
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
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objForm);
                objForm = null;
                GC.Collect();
            }
        }
    }
}
