using System;
using log4net;
using System.Runtime.InteropServices;

namespace T1.B1.Expenses
{
    public class Operations
    {
        private static Operations objExpenses;
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        

        private Operations()
        {

        }

        public static void formDataAddEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool blBubbleEvent)
        {
            if (objExpenses == null)
            {
                objExpenses = new Operations();
            }


            try
            {
                #region Legalizaciones
                #region Solicitud
                if (
                    BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.FormTypeEx == Settings._Main.ExpenseRequestUDoFormType &&
                    !BusinessObjectInfo.BeforeAction &&
                    BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    )
                {
                    


                }

                if (
                    BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.FormTypeEx == Settings._Main.ExpenseRequestUDoFormType &&
                    !BusinessObjectInfo.BeforeAction &&
                    BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    )
                {

                   


                }
                if (
                    BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.FormTypeEx == Settings._Main.ExpenseRequestUDoFormType &&
                    !BusinessObjectInfo.BeforeAction &&
                    BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    )
                {
                   





                }
                #endregion

                #region Legalizacion
                if (
                    BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.FormTypeEx == Settings._Main.LegalizationRequestUDoFormType &&
                    !BusinessObjectInfo.BeforeAction &&
                    BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    )
                {
                    Expenses.getLegalizationDocEntryOnLoad(BusinessObjectInfo);





                }

                if (
                    BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.FormTypeEx == Settings._Main.LegalizationRequestUDoFormType &&
                    !BusinessObjectInfo.BeforeAction &&
                    BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    )
                {
                    Expenses.changeRequestStatus(BusinessObjectInfo);





                }

                
                #endregion
                #endregion

                #region Caja Menor Legalizacion
                if (
                    BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.FormTypeEx == Settings._MainPettyCash.pettyCashLegalizationFormType &&
                    !BusinessObjectInfo.BeforeAction &&
                    BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    )
                {
                    PettyCash.getPCLegalizationDocEntryOnLoad(BusinessObjectInfo);
                }
                #endregion



            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("FormDataEvent Error", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("FormDataEvent Error", er);

            }
        }
        public static void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
       
            if(objExpenses == null)
            {
                objExpenses = new Operations();
                
            }
            try
            {
                #region Clasificación Tipos
                if (pVal.MenuUID == "HCO_MCLM009"
                    && !pVal.BeforeAction)
                {
                    MainObject.Instance.B1Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, Settings._Main.ExpenseTypeClasificatioonUDOFormType, "");
                }
                #endregion

                #region Conceptos
                if (pVal.MenuUID == "HCO_MCLM010"
                    && !pVal.BeforeAction)
                {
                    Expenses.openConceptsForm(pVal);
                }

                #endregion

                #region Tipos
                if (pVal.MenuUID == "HCO_MCLM011"
                    && !pVal.BeforeAction)
                {
                    Expenses.openExpenseTypeForm(pVal);
                }
                #endregion

                #region Solicitud
                if (pVal.MenuUID == "HCO_MCLM012"
                    && !pVal.BeforeAction)
                {
                    Expenses.openExpenseRequestForm(pVal);
                }

                if (pVal.MenuUID == "1282"
                    && !pVal.BeforeAction)
                {
                    SAPbouiCOM.Form objForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                    if(objForm.TypeEx == Settings._Main.ExpenseRequestUDoFormType)
                    {
                        Expenses.configureExpenseRequestFOrm(objForm.UniqueID, false);
                    }
                }
                #endregion

                #region Aprobación
                if (pVal.MenuUID == "HCO_MCLM013"
                    && !pVal.BeforeAction)
                {
                    Expenses.loadPendingAppovedRequestForm();
                }
                #endregion

                #region Desembolsos
                if (pVal.MenuUID == "HCO_MCLM014"
                    && !pVal.BeforeAction)
                {
                    Expenses.loadPaymentForm();
                }
                #endregion

                #region Legalización
                if (pVal.MenuUID == "HCO_MCLM015"
                    && !pVal.BeforeAction)
                {
                    Expenses.openLegalizationForm(pVal);
                }

                if (pVal.MenuUID == "1282"
                    && !pVal.BeforeAction)
                {
                    SAPbouiCOM.Form objForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                    if (objForm.TypeEx == Settings._Main.ExpenseRequestUDoFormType)
                    {
                        Expenses.configureLegalizationForm(objForm.UniqueID, false);
                    }
                }

                if (pVal.MenuUID == "HCO_T1EXP400_Remove_Line"
                    && !pVal.BeforeAction)
                {
                    string strLastUID = CacheManager.CacheManager.Instance.getFromCache(Settings._Main.LastExpenseActiveForm) == null ? "" : CacheManager.CacheManager.Instance.getFromCache(Settings._Main.LastExpenseActiveForm);
                    Expenses.deleteRowAfter(strLastUID);
                }
                #endregion

                #region Cajas Menores
                if (pVal.MenuUID == "HCO_MCLM003" //Conceptos
                    && !pVal.BeforeAction)
                {
                    MainObject.Instance.B1Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "HCO_T1PTC200", "");
                }

                if (pVal.MenuUID == "HCO_MCLM004" //Apertura
                    && !pVal.BeforeAction)
                {
                    PettyCash.loadPettyCashPaymentForm();
                }

                if (pVal.MenuUID == "HCO_MCLM005" //Creación
                    && !pVal.BeforeAction)
                {
                    MainObject.Instance.B1Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "HCO_T1PTC100", "");
                }

                if (pVal.MenuUID == "HCO_MCLM007" // Registro
                    && !pVal.BeforeAction)
                {
                    PettyCash.openLegalizationForm(pVal);
                }
    //            if (pVal.MenuUID == "HCO_MCLM007"
    //&& !pVal.BeforeAction)
    //            {
    //                MainObject.Instance.B1Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "HCO_T1PTC100", "");
    //                //MainObject.Instance.B1Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "HCO_T1PTC300", "");

    //            }

                if (pVal.MenuUID == "1282"
                    && !pVal.BeforeAction)
                {
                    SAPbouiCOM.Form objForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                    if (objForm.TypeEx == Settings._MainPettyCash.pettyCashPaymentFormType)
                    {
                        PettyCash.configureLegalizationForm(objForm.UniqueID, false);
                    }
                }

                #endregion




            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }        
        public static void RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        {
            try
            {
                if (eventInfo.EventType == SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK
                    && eventInfo.ItemUID == "0_U_G"
                    && eventInfo.BeforeAction

                        )
                {
                    CacheManager.CacheManager.Instance.addToCache("RightClickLastRow", eventInfo.FormUID + "#" + eventInfo.Row, CacheManager.CacheManager.objCachePriority.Default);


                }
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
        }
        public static void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {

            try
            {

                #region General

                CacheManager.CacheManager.Instance.addToCache(Settings._Main.LastExpenseActiveForm, pVal.FormUID, CacheManager.CacheManager.objCachePriority.NotRemovable);


                
                #endregion


                #region Conceptos
                #region Concepts ChooseFromList BP

                if (pVal.FormTypeEx == Settings._Main.ConceptUDOFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && pVal.ItemUID == "0_U_G"
                    && !pVal.BeforeAction
                    )
                {

                    SAPbouiCOM.Form objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                    Expenses.loadBPNameCFL(objForm, pVal);
                    
                }
                if (pVal.FormTypeEx == Settings._Main.ConceptUDOFormType && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "1" && pVal.BeforeAction)
                {

                    SAPbouiCOM.Form objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)objForm.Items.Item("0_U_G").Specific;
                    oMatrix.FlushToDataSource();

                }






                #endregion

                #region Filter Account CFL
                if (pVal.FormTypeEx == Settings._Main.ConceptUDOFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && pVal.BeforeAction
                    && pVal.ItemUID == "14_U_E"
                        )
                {
                    Expenses.filterAccountConceptsUDO(pVal, false);

                }

                if (pVal.FormTypeEx == Settings._Main.ConceptUDOFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "14_U_E"
                        )
                {
                    Expenses.filterAccountConceptsUDO(pVal, true);

                }
                #endregion

                #endregion

                #region Tipos
                

                #region Filter Account CFL
                if (pVal.FormTypeEx == Settings._Main.ExpenseTypeUDOFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && pVal.BeforeAction
                    && pVal.ItemUID == "16_U_E"
                        )
                {
                    Expenses.filterAccountExpTypeUDO(pVal, false);

                }

                if (pVal.FormTypeEx == Settings._Main.ExpenseTypeUDOFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "16_U_E"
                        )
                {
                    Expenses.filterAccountExpTypeUDO(pVal, true);

                }
                #endregion

                #endregion

                #region Solicitud
                if (pVal.FormTypeEx == Settings._Main.ExpenseRequestUDoFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    && !pVal.BeforeAction
                    
                        )
                {
                    CacheManager.CacheManager.Instance.addToCache(Settings._Main.ExpenseRequestFormLastId, pVal.FormUID, CacheManager.CacheManager.objCachePriority.Default);

                }

                if (pVal.FormTypeEx == Settings._Main.ExpenseRequestUDoFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && pVal.BeforeAction
                    && pVal.ItemUID == "1_U_G"
                   
                        )
                {
                    Expenses.filterStepTypeUDO(pVal, false);

                }
                if (pVal.FormTypeEx == Settings._Main.ExpenseRequestUDoFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "1_U_G"

                        )
                {
                    Expenses.filterStepTypeUDO(pVal, true);

                }

                if (pVal.FormTypeEx == Settings._Main.ExpenseRequestUDoFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && pVal.BeforeAction
                    && pVal.ItemUID == "26_U_E"

                        )
                {
                    Expenses.filterStatusUDO(pVal, false);

                }
                if (pVal.FormTypeEx == Settings._Main.ExpenseRequestUDoFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "26_U_E"

                        )
                {
                    Expenses.filterStatusUDO(pVal, true);

                }

                #endregion

                #region Aprobacion
                //if(pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                //    && pVal.FormTypeEx == "HCO_REQAPR"
                //    && pVal.ItemUID == "aprGrid"
                //    && pVal.ColUID == "Documento"
                //    && pVal.BeforeAction)
                //    {
                //    BubbleEvent = false;
                //    SAPbouiCOM.Form objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                //    SAPbouiCOM.DataTable objData = objForm.DataSources.DataTables.Item("HCO_RELIST");
                //    string strValue = Convert.ToString(objData.GetValue("Documento", pVal.Row));
                //    MainObject.Instance.B1Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "HCO_T1EXP600", strValue);
                //}

                if(pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && pVal.FormTypeEx == "HCO_REQAPR"
                    && pVal.ItemUID =="btnUpdate"
                    && !pVal.BeforeAction)
                {
                    Expenses.updateRequestStatus(pVal);
                }
                #endregion

                #region Desembolsos

                if (pVal.FormTypeEx == Settings._Main.PaymentFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && pVal.BeforeAction
                    && pVal.ItemUID == "txtAcct"
                        )
                {
                    Expenses.filterAccountPaymentForm(pVal, false);

                }

                if (pVal.FormTypeEx == Settings._Main.PaymentFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "txtAcct"
                        )
                {
                    Expenses.filterAccountPaymentForm(pVal, true);

                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && pVal.FormTypeEx == Settings._Main.PaymentFormType
                    && pVal.ItemUID == "btnCreate"
                    && !pVal.BeforeAction)
                {
                    Expenses.addPaymentDocument(pVal);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.FormTypeEx == Settings._Main.PaymentFormType && pVal.ItemUID == "gridDocs" && pVal.ColUID.Equals("Desembolsar") && !pVal.BeforeAction)
                {
                    Expenses.GetTotal_Desembolso(pVal);
                }

                #endregion

                #region Legalizacion

                if (pVal.FormTypeEx == Settings._Main.LegalizationRequestUDoFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    && !pVal.BeforeAction

                        )
                {
                    CacheManager.CacheManager.Instance.addToCache(Settings._Main.LegalizationFormLastId, pVal.FormUID, CacheManager.CacheManager.objCachePriority.Default);
                    

                }

                




                if (pVal.FormTypeEx == Settings._Main.LegalizationRequestUDoFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && pVal.BeforeAction
                    && pVal.ItemUID == "22_U_E"
                        )
                {
                    Expenses.filterValidRequestUDO(pVal, false);

                }

                if (pVal.FormTypeEx == Settings._Main.LegalizationRequestUDoFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "22_U_E"
                        )
                {
                    Expenses.filterValidRequestUDO(pVal, true);

                }

                if (pVal.FormTypeEx == Settings._Main.LegalizationRequestUDoFormType
                   && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                   && pVal.BeforeAction
                   && pVal.ItemUID == "0_U_G"
                   && pVal.ColUID == "C_0_1"

                       )
                {
                    Expenses.filterValidConceptsUDO(pVal, false, out BubbleEvent);

                }

                if (pVal.FormTypeEx == Settings._Main.LegalizationRequestUDoFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "0_U_G"
                    && pVal.ColUID == "C_0_1"
                        )
                {
                    Expenses.filterValidConceptsUDO(pVal, true, out BubbleEvent);

                }

                if (pVal.FormTypeEx == Settings._Main.LegalizationRequestUDoFormType
                   && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                   && pVal.BeforeAction
                   && pVal.ItemUID == "0_U_G"
                   && pVal.ColUID == "C_0_5"

                       )
                {
                    Expenses.filterThirdPartyConceptsUDO(pVal, false, out BubbleEvent);

                }

                if (pVal.FormTypeEx == Settings._Main.LegalizationRequestUDoFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "0_U_G"
                    && pVal.ColUID == "C_0_5"
                        )
                {
                    //Expenses.filterThirdPartyConceptsUDO(pVal, true, out BubbleEvent);
                    SAPbouiCOM.Form oForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.Matrix oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("0_U_G").Specific;
                    oMtx.FlushToDataSource();


                }

                if (pVal.FormTypeEx == Settings._Main.LegalizationRequestUDoFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "0_U_G"
                    //&& pVal.ColUID == "C_0_2"
                        )
                {
                    Tuple<string, string, int> tLastGotFocusColumn = new Tuple<string, string, int>(pVal.ItemUID, pVal.ColUID, pVal.Row);
                    CacheManager.CacheManager.Instance.addToCache(pVal.FormUID + "_LastColumnFocus", tLastGotFocusColumn, CacheManager.CacheManager.objCachePriority.Default);
                    //MainObject.Instance.B1Application.SetStatusBarMessage(pVal.Row.ToString(),SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    //Expenses.refreshLineInfo(pVal);

                }

                if (pVal.FormTypeEx == Settings._Main.LegalizationRequestUDoFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "0_U_G"
                    
                        )
                {
                    Tuple<string, string, int> tLastGotFocusColumn = CacheManager.CacheManager.Instance.getFromCache(pVal.FormUID + "_LastColumnFocus");
                    if (tLastGotFocusColumn != null)
                    {
                        if (tLastGotFocusColumn.Item1 == "0_U_G" && tLastGotFocusColumn.Item2 == "C_0_3")
                        {

                            Expenses.refreshFormValues(pVal, tLastGotFocusColumn);
                        }
                    }

                }

                if (pVal.FormTypeEx == Settings._Main.LegalizationRequestUDoFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "btnContab"

                        )
                {
                    Expenses.postLegalization(pVal);
                }



                #endregion

                #region Caja Menor

                #region Conceptos

                #region Concepts ChooseFromList BP

                if (pVal.FormTypeEx == Settings._MainPettyCash.pettyCashConceptUDOFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && pVal.ItemUID == "1_U_G"
                    && !pVal.BeforeAction
                    )
                {

                    SAPbouiCOM.Form objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                    PettyCash.loadBPNameCFL(objForm, pVal);

                }

                
                #endregion

                #region Filter Account CFL
                if (pVal.FormTypeEx == Settings._MainPettyCash.pettyCashConceptUDOFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && pVal.BeforeAction
                    && pVal.ItemUID == "14_U_E"
                        )
                {
                    PettyCash.filterAccountConceptsUDO(pVal, false);

                }

                if (pVal.FormTypeEx == Settings._MainPettyCash.pettyCashConceptUDOFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "14_U_E"
                        )
                {
                    PettyCash.filterAccountConceptsUDO(pVal, true);

                }
                #endregion
                #endregion

                #region Apertura

                if (pVal.FormTypeEx == Settings._MainPettyCash.pettyCashPaymentFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && pVal.BeforeAction
                    && pVal.ItemUID == "Item_7"
                        )
                {
                    PettyCash.filterAccountPCPaymentForm(pVal, false);

                }

                if (pVal.FormTypeEx == Settings._MainPettyCash.pettyCashPaymentFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "Item_7"
                        )
                {
                    PettyCash.filterAccountPCPaymentForm(pVal, true);

                }

                if (pVal.FormTypeEx == Settings._MainPettyCash.pettyCashPaymentFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && pVal.BeforeAction
                    && pVal.ItemUID == "Item_8"
                        )
                {
                    PettyCash.filterPettyCashPCPaymentForm(pVal, false);

                }

                if (pVal.FormTypeEx == Settings._MainPettyCash.pettyCashPaymentFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "Item_8"
                        )
                {
                    PettyCash.filterPettyCashPCPaymentForm(pVal, true);

                }

                if (pVal.FormTypeEx == Settings._MainPettyCash.pettyCashPaymentFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "btnOpen"
                        )
                {
                    PettyCash.addPaymentDocument(pVal);

                }
                #endregion

                #region Legalizacion

                if (pVal.FormTypeEx == Settings._MainPettyCash.pettyCashLegalizationFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    && !pVal.BeforeAction

                        )
                {
                    CacheManager.CacheManager.Instance.addToCache(Settings._Main.LegalizationFormLastId, pVal.FormUID, CacheManager.CacheManager.objCachePriority.Default);


                }

                if (pVal.FormTypeEx == Settings._MainPettyCash.pettyCashLegalizationFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && pVal.BeforeAction
                    && pVal.ItemUID == "21_U_E"
                        )
                {
                    PettyCash.filterValidPCUDO(pVal, false);

                }

                if (pVal.FormTypeEx == Settings._MainPettyCash.pettyCashLegalizationFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "21_U_E"
                        )
                {
                    PettyCash.filterValidPCUDO(pVal, true);

                }

                if (pVal.FormTypeEx == Settings._MainPettyCash.pettyCashLegalizationFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && pVal.ItemUID == "0_U_G"
                    && pVal.ColUID == "C_0_1"
                        )
                {
                    PettyCash.setValidConceptInfo(pVal,out BubbleEvent);

                }

                

                if (pVal.FormTypeEx == Settings._MainPettyCash.pettyCashLegalizationFormType
                   && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                   && pVal.BeforeAction
                   && pVal.ItemUID == "0_U_G"
                   && pVal.ColUID == "C_0_5"

                       )
                {
                    PettyCash.filterThirdPartyConceptsUDO(pVal, false, out BubbleEvent);

                }

                if (pVal.FormTypeEx == Settings._MainPettyCash.pettyCashLegalizationFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "0_U_G"
                    && pVal.ColUID == "C_0_5"
                        )
                {
                    PettyCash.filterThirdPartyConceptsUDO(pVal, true, out BubbleEvent);

                }

                if (pVal.FormTypeEx == Settings._MainPettyCash.pettyCashLegalizationFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "0_U_G"
                        //&& pVal.ColUID == "C_0_2"
                        )
                {
                    Tuple<string, string> tLastGotFocusColumn = new Tuple<string, string>(pVal.ItemUID, pVal.ColUID);
                    CacheManager.CacheManager.Instance.addToCache(pVal.FormUID + "_LastColumnFocus", tLastGotFocusColumn, CacheManager.CacheManager.objCachePriority.Default);
                    //Expenses.refreshLineInfo(pVal);

                }

                if (pVal.FormTypeEx == Settings._MainPettyCash.pettyCashLegalizationFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "0_U_G"

                        )
                {
                    Tuple<string, string> tLastGotFocusColumn = CacheManager.CacheManager.Instance.getFromCache(pVal.FormUID + "_LastColumnFocus");
                    if (tLastGotFocusColumn != null)
                    {
                        if (tLastGotFocusColumn.Item1 == "0_U_G" && tLastGotFocusColumn.Item2 == "C_0_3")
                        {

                            PettyCash.refreshFormValues(pVal);
                        }
                    }

                }

                if (pVal.FormTypeEx == Settings._MainPettyCash.pettyCashLegalizationFormType
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "btnContab"

                        )
                {
                   PettyCash.postLegalization(pVal);
                }


                #endregion




                #endregion


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



    }
}
