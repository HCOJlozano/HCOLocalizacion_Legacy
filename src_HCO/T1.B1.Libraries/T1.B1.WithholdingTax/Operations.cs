using System;
using System.Collections.Generic;
using log4net;
using Newtonsoft.Json;
using System.Runtime.InteropServices;
using SAPbouiCOM;

namespace T1.B1.WithholdingTax
{
    public enum FORM_MODE { SEARCH, NEW, OK }
    public class Operations
    {
        private static Operations objWithHoldingTax;
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static bool runResizelogic = true;
        private static List<string> WTDocuments = new List<string>();
        private static Form objForm;


        public static bool CloseFormBP = false;
        private Operations()
        {
        }
        public static void FormDataEvent(ref BusinessObjectInfo BusinessObjectInfo, ref bool blBubbleEvent)
        {
            if (objWithHoldingTax == null) objWithHoldingTax = new Operations();
            WTDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._WithHoldingTax.WTFormTypes);

            try
            {
                if (!BusinessObjectInfo.BeforeAction)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case BoEventTypes.et_FORM_DATA_ADD:
                            if (BusinessObjectInfo.ActionSuccess && (WTDocuments.Contains(BusinessObjectInfo.FormTypeEx))) WithholdingTax.AddDocumentInfo(BusinessObjectInfo);
                            break;
                        case BoEventTypes.et_FORM_UNLOAD:
                            if (BusinessObjectInfo.ActionSuccess && (WTDocuments.Contains(BusinessObjectInfo.FormTypeEx))) WithholdingTax.RemoveFromCache(BusinessObjectInfo.FormUID);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            if (BusinessObjectInfo.FormTypeEx == "HCO_FWT1100") WithholdingTax.InitBusinessPartnerForm(BusinessObjectInfo.FormUID);
                            break;
                    }
                }
                else
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case BoEventTypes.et_FORM_DATA_ADD:
                            if (WTDocuments.Contains(BusinessObjectInfo.FormTypeEx))
                            {
                                objForm = MainObject.Instance.B1Application.Forms.Item(BusinessObjectInfo.FormUID);
                                if (!WithholdingTax.HasRelParty(objForm.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0)))
                                {
                                    MainObject.Instance.B1Application.MessageBox("El Cliente/Proveedor no tiene tercero relacionado.");
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
            finally
            {
                if( objForm != null )
                 System.Runtime.InteropServices.Marshal.ReleaseComObject(objForm);

                GC.Collect();
            }
        }
        public static void MenuEvent(ref MenuEvent pVal, ref bool BubbleEvent)
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
                            WithholdingTax.InitWithHoldingOperForm(WithholdingTax.LoadWithHoldingForm(Settings.EWithHoldingTax.WITHHOLDING_OPERATION));
                            break;
                        case "HCO_MWTARU":
                            var eventInfoMrparu = CacheManager.CacheManager.Instance.getFromCache(Settings._Main.lastRightClickEventInfo);
                            B1.Base.UIOperations.FormsOperations.MatrixOperationUDO("Add", "0_U_G", MainObject.Instance.B1Application.Forms.ActiveForm);
                            break;
                        case "HCO_MWT0004":
                            WithholdingTax.InitMissingOperationsForm(WithholdingTax.LoadWithHoldingForm(Settings.EWithHoldingTax.MISSING_OPERATIONS));
                            break;
                        case "1281":
                            objForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                            if (objForm.TypeEx == "HCO_FWT1100") WithholdingTax.SetFormState(objForm, FORM_MODE.SEARCH);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objForm);
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
            finally
            {

                GC.Collect();
            }
        }
        public static void ItemEvent(string FormUID, ref ItemEvent pVal, ref bool BubbleEvent)
        {
            WTDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._WithHoldingTax.WTFormTypes);

            try
            {
                if (!pVal.BeforeAction)
                {
                    switch (pVal.EventType)
                    {
                        case BoEventTypes.et_CHOOSE_FROM_LIST:
                            if (pVal.FormTypeEx == "HCO_FWT0100") WithholdingTax.SetChooseFromListMunMatrix(pVal);
                            break;
                        case BoEventTypes.et_LOST_FOCUS:
                            if (WTDocuments.Contains(pVal.FormTypeEx))
                            {
                                if (pVal.Action_Success)
                                {
                                    objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                                    if ((pVal.ItemUID == "4" || pVal.ItemUID == "38") && objForm.Mode == BoFormMode.fm_ADD_MODE)
                                    {
                                        CacheManager.CacheManager.Instance.addToCache("WTLastActiveForm", pVal.FormUID, CacheManager.CacheManager.objCachePriority.Default);
                                        if (WithholdingTax.GetSelectedBPInformation(objForm, false)) WithholdingTax.activateWTMenu(pVal.FormUID, true);
                                    }
                                }
                            }
                            if (pVal.FormTypeEx.Equals("60504"))
                            {
                                string strLastActiveForm = CacheManager.CacheManager.Instance.getFromCache("WTLastActiveForm") == null ? "" : CacheManager.CacheManager.Instance.getFromCache("WTLastActiveForm");

                                if (pVal.ItemUID == "6" && !(CacheManager.CacheManager.Instance.getFromCache("Updating_" + strLastActiveForm) == null ? false : true) && (pVal.ColUID == "14" || pVal.ColUID == "U_HCO_BaseAmnt"))
                                {
                                    string blAutoActivate = CacheManager.CacheManager.Instance.getFromCache("WTAutoActivate") != null ? CacheManager.CacheManager.Instance.getFromCache("WTAutoActivate") : "";
                                    objForm = null;

                                    if (strLastActiveForm.Trim().Length > 0 || blAutoActivate.Trim().Length == 0)
                                    {
                                        bool isDisabled = CacheManager.CacheManager.Instance.getFromCache("Disable_" + strLastActiveForm) == null ? false : true;
                                        objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                                        if (objForm.Mode == BoFormMode.fm_UPDATE_MODE && !isDisabled)
                                        {
                                            if (MainObject.Instance.B1Application.MessageBox("La modificación manual de las retenciones deshabilitará el cálculo automático para este documento. ¿Desea Continuar? ", 2, "Sí", "No", "") != 2)
                                            {
                                                //if (strLastActiveForm.Trim().Length > 0)
                                                //{
                                                CacheManager.CacheManager.Instance.addToCache(string.Concat("Disable_", strLastActiveForm), true, CacheManager.CacheManager.objCachePriority.Default);
                                                //}
                                                //else
                                                //{
                                                //    if (objForm.Mode == BoFormMode.fm_UPDATE_MODE)
                                                //    {
                                                //        objForm.Items.Item("1").Click(BoCellClickType.ct_Regular);
                                                //    }
                                                //    else BubbleEvent = false;
                                                //}

                                                //objForm.Close();
                                            }
                                            //else
                                            //{
                                            //    BubbleEvent = false;
                                            //    objForm.Items.Item("2").Click(BoCellClickType.ct_Regular);
                                            //}
                                        }
                                        CacheManager.CacheManager.Instance.removeFromCache("WTLastActiveForm");
                                        //T1.B1.Base.UIOperations.Operations.stopProgressBar();


                                    }
                                }
                            }
                            break;
                        case BoEventTypes.et_COMBO_SELECT:
                            if (WTDocuments.Contains(pVal.FormTypeEx))
                            {
                                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                                if (pVal.ItemUID.Equals("226") && objForm.Mode == BoFormMode.fm_ADD_MODE)
                                {
                                    if (WithholdingTax.GetSelectedBPInformation(objForm, true)) WithholdingTax.activateWTMenu(pVal.FormUID, true);
                                }
                            }
                            break;

                        case BoEventTypes.et_ITEM_PRESSED:
                            
                            break;
                        case BoEventTypes.et_FORM_LOAD:
                            //if (WTDocuments.Contains(pVal.FormTypeEx))
                            //{
                            //    objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                            //    objForm.Items.Add("HCO_BTWT", BoFormItemTypes.it_BUTTON);
                            //    Button objBtn = (Button)objForm.Items.Item("HCO_BTWT").Specific;
                            //    objBtn.Caption = "Calcular retenciones";
                            //    objBtn.Item.Width = 110;
                            //    objBtn.Item.Top = objForm.Items.Item("2").Top;
                            //    objBtn.Item.Left = objForm.Items.Item("2").Left + objForm.Items.Item("2").Width + 2;
                            //    //objBtn.Item.Enabled = false;
                            //    objBtn.Item.Visible = true;
                            //}

                            if (pVal.FormTypeEx.Equals("60504"))
                            {
                                string strLastActiveForm = CacheManager.CacheManager.Instance.getFromCache("WTLastActiveForm") == null ? "" : CacheManager.CacheManager.Instance.getFromCache("WTLastActiveForm");
                                if (strLastActiveForm.Trim().Length > 0)
                                {
                                    bool blDisabled = CacheManager.CacheManager.Instance.getFromCache("Disable_" + strLastActiveForm) != null ? true : false;
                                    if (!blDisabled) WithholdingTax.SetTypeWT(pVal.FormUID, strLastActiveForm);
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
                            if (WTDocuments.Contains(pVal.FormTypeEx)) WithholdingTax.RemoveFromCache(pVal.FormUID);
                            break;
                    }
                }
                else
                {
                    switch (pVal.EventType)
                    {
                        case BoEventTypes.et_ITEM_PRESSED:
                            if (pVal.FormTypeEx.Equals("HCO_FWT1200"))
                            {
                                if (pVal.ItemUID == "btnAdd")
                                {
                                    objForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                                    if (objForm.PaneLevel == 1)
                                        WithholdingTax.CreateMissingOperations(pVal);
                                    else
                                        objForm.Close();
                                }
                            }
                            if (WTDocuments.Contains(pVal.FormTypeEx))
                            {
                                if (pVal.Action_Success)
                                {
                                    objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                                    if (pVal.ItemUID == "1" && objForm.Mode == BoFormMode.fm_ADD_MODE)
                                    {
                                        CacheManager.CacheManager.Instance.addToCache("WTLastActiveForm", pVal.FormUID, CacheManager.CacheManager.objCachePriority.Default);
                                        if (WithholdingTax.GetSelectedBPInformation(objForm, false)) WithholdingTax.activateWTMenu(pVal.FormUID, true);
                                    }
                                }
                            }

                            if (WTDocuments.Contains(pVal.FormTypeEx))
                            {
                                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);

                                if (pVal.ItemUID.Equals("173") && objForm.Mode == BoFormMode.fm_ADD_MODE)
                                {
                                    CacheManager.CacheManager.Instance.addToCache("WTLastActiveForm", pVal.FormUID, CacheManager.CacheManager.objCachePriority.Default);
                                    WithholdingTax.GetSelectedBPInformation(objForm, true);
                                }

                                //if (pVal.ItemUID.Equals("HCO_BTWT") && objForm.Mode == BoFormMode.fm_ADD_MODE)
                                //{
                                //    objForm.Update();
                                //    WithholdingTax.GetSelectedBPInformation(objForm);
                                //    WithholdingTax.activateWTMenu(pVal.FormUID, true);
                                //}
                            }

                            break;
                        case BoEventTypes.et_LOST_FOCUS:


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
                GC.Collect();
                T1.B1.Base.UIOperations.Operations.stopProgressBar();
            }
        }
        public static void RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        {
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
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objForm);
                GC.Collect();
            }
        }

    }
}
