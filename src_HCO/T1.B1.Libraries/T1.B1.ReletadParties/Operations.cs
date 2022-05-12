using System;
using log4net;
using System.Runtime.InteropServices;
using SAPbouiCOM;
using System.Reflection;
using System.Xml;
using System.Threading;
using System.Security.Permissions;
using System.IO;

namespace T1.B1.RelatedParties
{
    public class Operations
    {
        private static Operations objReletadParties;
        public static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static bool runResizelogic = true;
        private static string RetFileName;
        public static bool VisibleWith = true;
        private Operations()
        {

        }

        public static void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            if (objReletadParties == null) objReletadParties = new Operations();

            try
            {
                if (pVal.BeforeAction)
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:

                            if (pVal.FormTypeEx == Settings._Main.BPFormTypeEx)
                            {
                                XmlDocument oXML = new XmlDocument();
                                oXML.Load(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Forms\\BP_MasterDataFrom.srf");
                                var oXml = oXML.InnerXml;
                                MainObject.Instance.B1Application.LoadBatchActions(string.Format(oXml, pVal.FormUID));
                                runResizelogic = false;
                            }

                            break;
                        case BoEventTypes.et_CLICK:
                            if (pVal.ItemUID == "Item_69")
                            {
                                BubbleEvent = Instance.ValidateFieldsMovement(pVal);
                            }
                            break;
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                            if (pVal.FormTypeEx == Settings._Main.RelatedPartiesUDO)
                            {
                                if (pVal.ItemUID == "1")
                                {
                                    BubbleEvent = Instance.ValidateFields(pVal);
                                    return;
                                }
                            }

                            if (pVal.FormTypeEx == Settings._Main.OutgoingPaymentFormTypeEx)
                            {
                                if (pVal.ItemUID == "1")
                                {
                                    BubbleEvent = Instance.ValidateFieldsPayment(pVal);
                                    return;
                                }
                            }

                            if (pVal.FormTypeEx == "940")
                            {
                                if (pVal.ItemUID == "1")
                                {
                                    BubbleEvent = Instance.ValidateFieldInventoryTransfer(pVal);
                                    return;
                                }
                            }

                            if (pVal.FormTypeEx.Equals("HCO_FRP0001"))
                            {
                                if (pVal.ItemUID == "1")
                                {
                                    Form oForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                                    if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                                        oForm.DataSources.DBDataSources.Item(0).SetValue("Code", 0, "1");
                                    return;
                                } 
                                else if(pVal.ItemUID == "Item_13")
                                {
                                    var third = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID).DataSources.DBDataSources.Item("@HCO_RP0001").GetValue("U_DefaultSN", 0);
                                    MainObject.Instance.B1Application.Menus.Item("HCO_MRP0009").Activate();
                                    var formThird = MainObject.Instance.B1Application.Forms.ActiveForm;
                                    MainObject.Instance.B1Application.Menus.Item("1281").Activate();
                                    ((EditText)formThird.Items.Item("0_U_E").Specific).Value = third;
                                    formThird.Items.Item("1").Click();                                    
                                }
                            }

                            if (pVal.FormTypeEx == Settings._Main.BPFormTypeEx)
                            {
                                if (pVal.ItemUID == "LinkRP")
                                {
                                    var third = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID).DataSources.UserDataSources.Item("UD_RelPty").Value;
                                    MainObject.Instance.B1Application.Menus.Item("HCO_MRP0009").Activate();
                                    var formThird = MainObject.Instance.B1Application.Forms.ActiveForm;
                                    MainObject.Instance.B1Application.Menus.Item("1281").Activate();
                                    ((EditText)formThird.Items.Item("0_U_E").Specific).Value = third;
                                    formThird.Items.Item("1").Click();
                                }
                            }

                            break;
                    }
                }
                else if (pVal.ActionSuccess)
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:

                            if (pVal.ActionSuccess)
                            {
                                if (pVal.FormTypeEx == Settings._Main.OutgoingPaymentFormTypeEx || 
                                    pVal.FormTypeEx == Settings._Main.ReceiptPaymentFormTypeEx || 
                                    pVal.FormTypeEx == Settings._Main.JournalFormTypeEx)
                                {
                                    if (pVal.ColUID == "U_HCO_RELPAR" || pVal.ItemUID == "txtRelpar")
                                        Instance.SetChooseFromListThirdPayment(pVal);
                                }
                                else
                                {
                                    if (pVal.ItemUID == "Item_52")
                                        Instance.SetChooseFromListMatrix(pVal);
                                    else
                                        Instance.SetChooseFromList(pVal);
                                }
                            }
                            break;

                        case BoEventTypes.et_COMBO_SELECT:
                            if (pVal.ActionSuccess)
                            {
                                if (pVal.ItemUID == "Item_74" || pVal.ItemUID == "Item_76")
                                    Instance.CheckLevelCondition(pVal);
                            }

                            break;

                        case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:

                            if (pVal.FormTypeEx == "HCO_FRP1100")
                            {
                                if (pVal.ActionSuccess)
                                {
                                    if (pVal.ItemUID == "Item_13")
                                        Instance.SetVerificationDigit(pVal.FormUID);
                                }
                                if (pVal.ItemUID == "Item_35")
                                {
                                    Form oform = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                                    if (oform.DataSources.DBDataSources.Item("@HCO_RP1100").GetValue("U_TaxCardTypeID", 0) == "1")
                                    {
                                        oform.Items.Item("Item_31").Enabled = true;
                                        oform.Items.Item("Item_32").Enabled = true;
                                        oform.Items.Item("Item_33").Enabled = true;
                                        oform.Items.Item("Item_39").Enabled = true;
                                        oform.Items.Item("Item_40").Enabled = true;
                                    }
                                    else
                                    {
                                        oform.Items.Item("Item_31").Enabled = false;
                                        oform.Items.Item("Item_32").Enabled = false;
                                        oform.Items.Item("Item_33").Enabled = false;
                                        oform.Items.Item("Item_39").Enabled = false;
                                        oform.Items.Item("Item_40").Enabled = false;
                                    }
                                }
                            }

                            break;

                        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:

                            if (pVal.FormTypeEx == Settings._Main.RelatedPartiesUDO)
                            {
                                Form oForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                                if (oForm.Mode == BoFormMode.fm_ADD_MODE) oForm.Items.Item("0_U_E").Enabled = true;
                            }

                            if (pVal.FormTypeEx == Settings._Main.OutgoingPaymentFormTypeEx ||
                                pVal.FormTypeEx == Settings._Main.ReceiptPaymentFormTypeEx ||
                                pVal.FormTypeEx == Settings._Main.JournalFormTypeEx)
                            {
                                Instance.LoadChooseFromListPayment(pVal);

                                if (pVal.FormTypeEx == Settings._Main.OutgoingPaymentFormTypeEx || pVal.FormTypeEx == Settings._Main.ReceiptPaymentFormTypeEx)
                                    Instance.LoadFieldPayment(pVal);
                            }

                            break;

                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                            if (pVal.FormTypeEx == Settings._Main.RelatedPartiesMovementReport)
                            {
                                if (pVal.ItemUID == "Item_69")
                                {
                                    Instance.OpenCrystalReport(pVal.FormUID);
                                    return;
                                }
                            }

                            if (pVal.FormTypeEx.Equals("HCO_FRP2100"))
                            {
                                Form oForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                                if (pVal.ItemUID.Equals("Item_2"))
                                {
                                    
                                    string archivo = OpenFileDialogForProcess();

                                    if (!archivo.Equals(string.Empty)) oForm.DataSources.UserDataSources.Item("UD_URL").Value = archivo;
                                    else oForm.DataSources.UserDataSources.Item("UD_URL").Value = "";
                                }
                                if (pVal.ItemUID.Equals("Item_3"))
                                {
                                    if (oForm.PaneLevel == 1)
                                    {
                                        if (oForm.DataSources.UserDataSources.Item("UD_URL").Value.ToString().Equals(string.Empty))
                                            MainObject.Instance.B1Application.MessageBox("Debe seleccionar un archivo.");
                                        else
                                        {
                                            oForm.Freeze(true);
                                            if (Instance.ProcessFile(oForm))
                                            {
                                                oForm.PaneLevel = 2;
                                                oForm.Width = 800;
                                                oForm.Height = 500;
                                                oForm.Items.Item("Item_3").Top = 422;
                                                oForm.Items.Item("Item_3").Left = 721;
                                                oForm.Items.Item("2").Top = 422;
                                                oForm.Items.Item("2").Left = 585;
                                            }
                                            else
                                            {
                                                MainObject.Instance.B1Application.MessageBox("Error leyendo archivo.");
                                            }

                                            oForm.Freeze(false);
                                        }
                                    }
                                    else if (oForm.PaneLevel == 2)
                                    {
                                        oForm.Freeze(true);
                                        Instance.createMissingRelatedParties(oForm);
                                        oForm.PaneLevel = 3;
                                        ((Button)oForm.Items.Item("Item_3").Specific).Caption = "Finalizar";
                                        oForm.Freeze(false);
                                    }
                                    else if (oForm.PaneLevel == 3)
                                    {
                                        oForm.Close();
                                    }

                                    }
                                }

                                if(pVal.FormTypeEx == "426" )
                                {
                                    if( pVal.ActionSuccess )
                                    {
                                        if( pVal.ItemUID == "arrowTer")
                                            Instance.OpenThirdForm(pVal);
                                    }
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

        private static string OpenFileDialogForProcess()
        {
            try
            {
                var oGetFileName = new GetFileNameClass
                {

                };

                Thread FileThread = new Thread(new ThreadStart(oGetFileName.GetFileName));
                FileThread.SetApartmentState(ApartmentState.STA);
                FileThread.Priority = ThreadPriority.Highest;
                FileThread.Start();

                while (!FileThread.IsAlive) ;
                Thread.Sleep(1);
                FileThread.Join();

                return oGetFileName.Path;
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


            return "";
        }

        private static void GetTheFile()
        {
            try
            {
                System.Windows.Forms.OpenFileDialog FileDialog = new System.Windows.Forms.OpenFileDialog();
                FileDialog.Multiselect = false;
                FileDialog.Filter = "Archivos TXT(*.txt)|*.txt";
                FileDialog.ShowDialog();
                RetFileName = FileDialog.FileName;
                System.Windows.Forms.Application.ExitThread();
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



        public static void FormDataAddEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool blBubbleEvent)
        {
            if (objReletadParties == null) objReletadParties = new Operations();

            try
            {
                if (BusinessObjectInfo.BeforeAction)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            if (BusinessObjectInfo.FormTypeEx == "HCO_FRP1100") Instance.CheckLinesUDO(BusinessObjectInfo);
                            break;
                    }

                }
                else if (BusinessObjectInfo.ActionSuccess)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            if (BusinessObjectInfo.FormTypeEx == "134") Instance.LoadDataThird(BusinessObjectInfo);
                            if (BusinessObjectInfo.FormTypeEx == "HCO_FRP1100")
                            {
                                Form oForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                                oForm.Items.Item("0_U_E").Enabled = false;
                            }
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:

                            if (BusinessObjectInfo.FormTypeEx == "426")
                            {
                                if (BusinessObjectInfo.ActionSuccess)
                                {
                                    Instance.UpdateJournalPayment(BusinessObjectInfo);
                                }
                            }

                            if (BusinessObjectInfo.FormTypeEx == "1470000009")
                            {
                                if (BusinessObjectInfo.ActionSuccess)
                                {
                                    var formActive = MainObject.Instance.B1Application.Forms.Item(BusinessObjectInfo.FormUID);
                                    var docEntry = formActive.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0);
                                    Instance.ActualizarInfoCapitalizacion(docEntry);
                                }
                            }

                            if (BusinessObjectInfo.FormTypeEx == "14700000037")
                            {
                                if (BusinessObjectInfo.ActionSuccess)
                                {
                                    var formActive = MainObject.Instance.B1Application.Forms.Item(BusinessObjectInfo.FormUID);
                                    var docEntry = formActive.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0);
                                    Instance.ActualizarInfoCapitalizacion(docEntry);
                                }
                            }

                            if (BusinessObjectInfo.FormTypeEx == "141")
                            {
                                if (BusinessObjectInfo.ActionSuccess)
                                {
                                    var formActive = MainObject.Instance.B1Application.Forms.Item(BusinessObjectInfo.FormUID);
                                    Instance.SetValueCapitalizacion(formActive.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0), BusinessObjectInfo.Type);
                                }
                            }
                            else if(BusinessObjectInfo.FormTypeEx == "181")
                            {
                                if (BusinessObjectInfo.ActionSuccess)
                                {
                                    var formActive = MainObject.Instance.B1Application.Forms.Item(BusinessObjectInfo.FormUID);
                                    Instance.SetCapitalizationNC(BusinessObjectInfo.FormUID, formActive.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0), BusinessObjectInfo.Type);
                                }
                            }

                            //Se actualiza registro de diario con el código del tercero, se filtra por lista de tipo de documentos.
                            if (Array.IndexOf(Parameters.docTercDef, BusinessObjectInfo.Type) >= 0 || Array.IndexOf(Parameters.docTercRel, BusinessObjectInfo.Type) >= 0)
                            {
                                if (BusinessObjectInfo.Type == "24" || BusinessObjectInfo.Type == "46")
                                    Instance.UpdateJournalPaymentCreated(BusinessObjectInfo.Type, BusinessObjectInfo.ObjectKey);
                                else
                                    Instance.UpdateJournalDocumentCreated(BusinessObjectInfo.Type, BusinessObjectInfo.ObjectKey);
                            }

                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            if (BusinessObjectInfo.FormTypeEx == "HCO_FRP1100") Instance.FrozenForRelParty(BusinessObjectInfo);
                            break;
                    }
                }
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
            if (objReletadParties == null) objReletadParties = new Operations();


            try
            {
                if (!pVal.BeforeAction)
                {
                    Form oForm = null;

                    switch (pVal.MenuUID)
                    {
                        case "48255":
                            Instance.LoadRelatedPartiesForm(Settings.RelatedParties.RELATED_PARTIES_MUNICIPALITY);
                            break;
                        case "HCO_MRP0001":
                            Instance.LoadRelatedPartiesForm(Settings.RelatedParties.RELATED_PARTIES_CONFIGURATION);
                            break;
                        case "HCO_MRP0002":
                            Instance.LoadRelatedPartiesForm(Settings.RelatedParties.RELATED_PARTIES_TYPES);
                            break;
                        case "HCO_MRP0003":
                            Instance.LoadRelatedPartiesForm(Settings.RelatedParties.RELATED_PARTIES_DOCUMENT_TYPES);
                            break;
                        case "HCO_MRP0004":
                            Instance.LoadRelatedPartiesForm(Settings.RelatedParties.RELATED_PARTIES_DEPARTMENT);
                            break;
                        case "HCO_MRP0005":
                            Instance.LoadRelatedPartiesForm(Settings.RelatedParties.RELATED_PARTIES_MUNICIPALITY);
                            break;
                        case "HCO_MRP0006":
                            Instance.LoadRelatedPartiesForm(Settings.RelatedParties.RELATED_PARTIES_TYPES_CONTRIB);
                            break;
                        case "HCO_MRP0007":
                            Instance.LoadRelatedPartiesForm(Settings.RelatedParties.RELATED_PARTIES_ECONOMIC_ACTIVITY);
                            break;
                        case "HCO_MRP0008":
                            Instance.LoadRelatedPartiesForm(Settings.RelatedParties.RELATED_PARTIES_TRIBUTARY_REGIMEN);
                            break;
                        case "HCO_MRP0009":
                            Instance.LoadRelatedPartiesForm(Settings.RelatedParties.RELATED_PARTIES);
                            break;
                        case "HCO_MRP1010":
                            Instance.LoadRelatedPartiesForm(Settings.RelatedParties.RELATED_PARTIES_MOVEMENT);
                            break;
                        case "HCO_MRP0010":
                            Instance.LoadRelatedPartiesForm(Settings.RelatedParties.RELATED_PARTIES_CREATION_WIZARD);
                            break;
                        case "HCO_RPT0011":
                            Instance.LoadCrystalReport(TYPE_CRYSTAL.ERI);
                            break;
                        case "HCO_RPT0012":
                            Instance.LoadCrystalReport(TYPE_CRYSTAL.ESFA);
                            break;
                        case "HCO_RPT0013":
                            Instance.LoadCrystalReport(TYPE_CRYSTAL.BALANCE);
                            break;
                        case "HCO_RPT0014":
                            Instance.LoadCrystalReport(TYPE_CRYSTAL.DIARIO);
                            break;
                        case "HCO_RPT0015":
                            Instance.LoadCrystalReport(TYPE_CRYSTAL.CERT_RET);
                            break;
                        case "HCO_RPT0016":
                            Instance.LoadCrystalReport(TYPE_CRYSTAL.AUXILIAR);
                            break;
                        case "HCO_RPT0017":
                            Instance.LoadCrystalReport(TYPE_CRYSTAL.RETPURCH_COD);
                            break;
                        case "HCO_RPT0018":
                            Instance.LoadCrystalReport(TYPE_CRYSTAL.RETSALE_COD);
                            break;
                        case "HCO_RPT0019":
                            Instance.LoadCrystalReport(TYPE_CRYSTAL.RETPRUCH_CARD);
                            break;
                        case "HCO_RPT0020":
                            Instance.LoadCrystalReport(TYPE_CRYSTAL.RETSALE_CARD);
                            break;
                        case "HCO_RPT0021":
                            Instance.LoadCrystalReport(TYPE_CRYSTAL.IVASALE_COD);
                            break;
                        case "HCO_RPT0022":
                            Instance.LoadCrystalReport(TYPE_CRYSTAL.IVAPRUCH_COD);
                            break;

                        //case "1282":
                        //    Instance.addLineToDS();
                        //    break;
                        case "HCO_MRPAR":
                            var eventInfoMrpar = CacheManager.CacheManager.Instance.getFromCache(Settings._Main.lastRightClickEventInfo);
                            Instance.relatedPartiedMatrixOperation(eventInfoMrpar, "Add");
                            break;
                        case "HCO_MRPDR":
                            var eventInfoMrdpr = CacheManager.CacheManager.Instance.getFromCache(Settings._Main.lastRightClickEventInfo);
                            Instance.relatedPartiedMatrixOperation(eventInfoMrdpr, "Delete");
                            break;
                        case "HCO_MRP03":
                            Instance.loadMissingRelatedPartiesForm();
                            break;
                        case "HCO_MRPARU":
                            var eventInfoMrparu = CacheManager.CacheManager.Instance.getFromCache(Settings._Main.lastRightClickEventInfo);
                            oForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                            Instance.RelatedPartiedMatrixOperationUDO("Add", oForm);
                            B1.Base.UIOperations.FormsOperations.UpdateMatrixRowNumbers("Item_52", oForm);
                            break;
                        case "HCO_MRPDRU":
                            var eventInfoMrpdru = CacheManager.CacheManager.Instance.getFromCache(Settings._Main.lastRightClickEventInfo);
                            var result = MainObject.Instance.B1Application.MessageBox("¿Qué desea realizar?", 1, "Desasociar", "Eliminar socio de negocio", "Cancelar");
                            oForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                            if (result == 1)
                            {
                                if (MainObject.Instance.B1Application.MessageBox("Esta acción causará que el cliente/proveedor no pueda ser utilizado en documentos y asientos contables, ¿Está seguro?", 2, "Si", "No", "Cancelar") == 1)
                                    Instance.RelatedPartiedMatrixOperationUDO("Delete", oForm);
                            }
                            else if (result == 2)
                            {
                                if (MainObject.Instance.B1Application.MessageBox("¿Está seguro que desea eliminar definitivamente el socio de negocio? Esta acción no se puede deshacer.", 2, "Si", "No", "Cancelar") == 1)
                                    Instance.DeleteRowBP();
                                MainObject.Instance.B1Application.ActivateMenuItem("1304");
                            }

                            break;
                        case "HCO_MRPACL":
                            Instance.createBP(TYPE_BP.CUSTOMER, Instance.GetNext_BPSecuence(TYPE_BP.CUSTOMER));
                            break;
                        case "HCO_MRPAPR":
                            Instance.createBP(TYPE_BP.SUPPLIER, Instance.GetNext_BPSecuence(TYPE_BP.SUPPLIER));
                            break;
                        case "HCO_MRPAEL":
                            Instance.DeleteThirdParty();
                            break;
                        case "HCO_SAPBP":
                            Instance.DeleteBP();
                            break;
                        case "1282":
                            oForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                            if (oForm.TypeEx == Settings._Main.RelatedPartiesUDO) oForm.Items.Item("0_U_E").Enabled = true;
                            break;
                        case "1281":
                            oForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                            if (oForm.TypeEx == Settings._Main.RelatedPartiesUDO) oForm.Items.Item("0_U_E").Enabled = true;
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

        public static void RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        {
            if (objReletadParties == null) objReletadParties = new Operations();

            SAPbouiCOM.Form objForm = null;
            
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(eventInfo.FormUID);

                if (eventInfo.BeforeAction)
                {
                    switch (objForm.TypeEx)
                    {
                        case "HCO_FRP1100":
                            if (eventInfo.ItemUID == "Item_52")
                            {
                                B1.Base.UIOperations.FormsOperations.AddRightClickMenu("HCO_MRPARU", "Agregar línea", 1);
                                B1.Base.UIOperations.FormsOperations.AddRightClickMenu("HCO_MRPDRU", "Eliminar línea", 2);
                                if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                {
                                    B1.Base.UIOperations.FormsOperations.AddRightClickMenu("HCO_MRPACL", "Agregar cliente", 3);
                                    B1.Base.UIOperations.FormsOperations.AddRightClickMenu("HCO_MRPAPR", "Agregar proveedor", 4);
                                }
                            }
                            else
                            {
                                DeleteThirdPartyRightClickMenus();
                            }
                            B1.Base.UIOperations.FormsOperations.AddRightClickMenu("HCO_MRPAEL", "Eliminar tercero", 5);
                            MainObject.Instance.B1Application.Menus.Item("1283").Enabled = false;
                            MainObject.Instance.B1Application.Menus.Item("1284").Enabled = false;
                            break;
                        default:
                            DeleteThirdPartyRightClickMenus();
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
        private static void DeleteThirdPartyRightClickMenus()
        {
            B1.Base.UIOperations.FormsOperations.DeleteRightClickMenu("HCO_MRPARU");
            B1.Base.UIOperations.FormsOperations.DeleteRightClickMenu("HCO_MRPDRU");
            B1.Base.UIOperations.FormsOperations.DeleteRightClickMenu("HCO_MRPACL");
            B1.Base.UIOperations.FormsOperations.DeleteRightClickMenu("HCO_MRPAPR");
            B1.Base.UIOperations.FormsOperations.DeleteRightClickMenu("HCO_MRPAEL");
        }
    }
}
