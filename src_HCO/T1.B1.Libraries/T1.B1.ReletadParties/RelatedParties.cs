using System;
using log4net;
using System.Runtime.InteropServices;
using System.Xml;
using System.Drawing;
using SAPbouiCOM;
using SAPbobsCOM;
using System.Reflection;
using T1.Queries.Entities;
using System.IO;
using System.Text;
using System.Data;
using System.Linq;
using System.Collections.Generic;

namespace T1.B1.RelatedParties
{
    public enum TYPE_BP { CUSTOMER, SUPPLIER };
    public enum TYPE_CRYSTAL { BALANCE, ERI, ESFA, DIARIO, TERCERO, RET_CODE, AUXILIAR, CERT_RET, RETPURCH_COD, RETSALE_COD, RETPRUCH_CARD, RETSALE_CARD, IVAPRUCH_COD, IVASALE_COD, BALNCE_TEST_RP, CERT_RET_IVA };


    public class Instance
    {
        private static Instance objRelParty;
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static List<string> listRow = new List<string>();

        private Instance()
        {
            if (objRelParty == null) objRelParty = new Instance();
            GetRelPartyConfiguration();
        }

        public static void GetRelPartyConfiguration()
        {
            Recordset oRS = (SAPbobsCOM.Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                oRS.DoQuery(Queries.Instance.Queries().Get("GetRelPartyConfig"));
                if (oRS.RecordCount > 0)
                {
                    CacheManager.CacheManager.Instance.addToCache("CodeRPConf", oRS.Fields.Item("Code").Value, CacheManager.CacheManager.objCachePriority.Default);
                    CacheManager.CacheManager.Instance.addToCache("DefaultSN", oRS.Fields.Item("U_DefaultSN").Value, CacheManager.CacheManager.objCachePriority.Default);
                    CacheManager.CacheManager.Instance.addToCache("ClientPrefix", oRS.Fields.Item("U_ClientPrefix").Value, CacheManager.CacheManager.objCachePriority.Default);
                    CacheManager.CacheManager.Instance.addToCache("VendorPrefix", oRS.Fields.Item("U_VendorPrefix").Value, CacheManager.CacheManager.objCachePriority.Default);
                    CacheManager.CacheManager.Instance.addToCache("CSeries", oRS.Fields.Item("U_CSeries").Value, CacheManager.CacheManager.objCachePriority.Default);
                    CacheManager.CacheManager.Instance.addToCache("VSeries", oRS.Fields.Item("U_VSeries").Value, CacheManager.CacheManager.objCachePriority.Default);
                    CacheManager.CacheManager.Instance.addToCache("ManBPSerie", oRS.Fields.Item("U_ManBPSerie").Value, CacheManager.CacheManager.objCachePriority.Default);
                    CacheManager.CacheManager.Instance.addToCache("MultBP", oRS.Fields.Item("U_MultBP").Value, CacheManager.CacheManager.objCachePriority.Default);
                    CacheManager.CacheManager.Instance.addToCache("AutoConse", oRS.Fields.Item("U_AutoConse").Value, CacheManager.CacheManager.objCachePriority.Default);
                    CacheManager.CacheManager.Instance.addToCache("TerPerfix", oRS.Fields.Item("U_TerPerfix").Value, CacheManager.CacheManager.objCachePriority.Default);
                    CacheManager.CacheManager.Instance.addToCache("NumChara", oRS.Fields.Item("U_NumChara").Value, CacheManager.CacheManager.objCachePriority.Default);
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
        private static void GetRelPartyName()
        {
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);


            oRS.DoQuery(string.Format(Queries.Instance.Queries().Get("CheckBPCant"), CacheManager.CacheManager.Instance.getFromCache("DefaultSN")));

        }

        public static void LoadChooseFromListPayment(ItemEvent pVal)
        {
            var form = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            var oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_ChooseFromListCreationParams);
            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "HCO_FRP1100";
            oCFLCreationParams.UniqueID = "CF_TER";

            form.ChooseFromLists.Add(oCFLCreationParams);

            var matrixItem = pVal.FormTypeEx == Settings._Main.JournalFormTypeEx ? "76" : "71";
            var matrix = (SAPbouiCOM.Matrix)form.Items.Item(matrixItem).Specific;
            matrix.Columns.Item("U_HCO_RELPAR").ChooseFromListUID = "CF_TER";
            matrix.Columns.Item("U_HCO_RELPAR").ChooseFromListAlias = "Code";
        }

        public static void SetChooseFromListContPlan(ItemEvent pVal)
        {
            var form = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            var lblTxt = form.Items.Add("lblThird", BoFormItemTypes.it_STATIC);
            lblTxt.Width = 150;
            lblTxt.Top = form.Items.Item("7").Top;
            lblTxt.Left = form.Items.Item("5").Left + form.Items.Item("5").Width - 130;
            ((StaticText)lblTxt.Specific).Caption = "Tercero Relacionado";

            var itms = form.Items.Add("txtThird", BoFormItemTypes.it_EDIT);
            itms.Top = form.Items.Item("5").Top;
            itms.Left = form.Items.Item("5").Left + form.Items.Item("5").Width - 130;

            ((EditText)itms.Specific).DataBind.SetBound(true, "OTRT", "U_HCO_RELPAR");
            ((EditText)itms.Specific).TabOrder = 9999;

            var oCFLs = form.ChooseFromLists;
            var oCFLCreationParams = ((ChooseFromListCreationParams)(MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_ChooseFromListCreationParams)));

            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "HCO_FRP1100";
            oCFLCreationParams.UniqueID = "CFLTHR";

            var oCFL = oCFLs.Add(oCFLCreationParams);

            ((EditText)form.Items.Item("txtThird").Specific).ChooseFromListUID = "CFLTHR";
            ((EditText)form.Items.Item("txtThird").Specific).ChooseFromListAlias = "Code";
        }

        public static void SetChooseFromListContPer(ItemEvent pVal)
        {
            var form = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            var lblTxt = form.Items.Add("lblThird", BoFormItemTypes.it_STATIC);
            lblTxt.Width = 150;
            lblTxt.Top = form.Items.Item("26").Top;
            lblTxt.Left = form.Items.Item("26").Left;
            ((StaticText)lblTxt.Specific).Caption = "Tercero Relacionado";

            var itms = form.Items.Add("txtThird", BoFormItemTypes.it_EDIT);
            itms.Top = form.Items.Item("22").Top;
            itms.Left = form.Items.Item("22").Left + 5;
            ((EditText)itms.Specific).DataBind.SetBound(true, "ORCR", "U_HCO_RELPAR");
            ((EditText)itms.Specific).TabOrder = 9999;

            var oCFLs = form.ChooseFromLists;
            var oCFLCreationParams = ((ChooseFromListCreationParams)(MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_ChooseFromListCreationParams)));

            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "HCO_FRP1100";
            oCFLCreationParams.UniqueID = "CFLTHR";

            var oCFL = oCFLs.Add(oCFLCreationParams);

            ((EditText)form.Items.Item("txtThird").Specific).ChooseFromListUID = "CFLTHR";
            ((EditText)form.Items.Item("txtThird").Specific).ChooseFromListAlias = "Code";
        }

        public static void SetContPer(ItemEvent pVal)
        {
            listRow.Clear();
            var form = MainObject.Instance.B1Application.Forms.ActiveForm;
            var matrix = (Matrix)form.Items.Item("3").Specific;
            for (int i = 1; i <= matrix.RowCount; i++)
            {
                if (matrix.IsRowSelected(i))
                    listRow.Add("'" + ((EditText)matrix.GetCellSpecific("1", i)).Value + "'");
            }
        }

        public static void MakeContPer(ItemEvent pVal)
        {
            if (listRow.Count > 0)
            {
                var journal = (JournalEntries)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oJournalEntries);
                var query = string.Format(Queries.Instance.Queries().Get("GetJournalContPer"), string.Join(",", listRow));
                var record = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                record.DoQuery(query);

                if (record.RecordCount > 0)
                {
                    while (!record.EoF)
                    {
                        if (journal.GetByKey(int.Parse(record.Fields.Item("TransId").Value.ToString())))
                        {
                            for (int i = 0; i < journal.Lines.Count; i++)
                            {
                                journal.Lines.SetCurrentLine(i);
                                journal.Lines.UserFields.Fields.Item("U_HCO_RELPAR").Value = record.Fields.Item("U_HCO_RELPAR").Value;
                            }

                            journal.Update();
                        }

                        record.MoveNext();
                    }
                }
            }
        }

        public static void LoadFieldPayment(ItemEvent pVal)
        {
            var form = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            var oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_ChooseFromListCreationParams);
            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "HCO_FRP1100";
            oCFLCreationParams.UniqueID = "CF_TERP";

            form.ChooseFromLists.Add(oCFLCreationParams);
            var refLbl = form.Items.Item("53");
            var refTxt = form.Items.Item("52");
            var linkTer = form.Items.Add("arrowTer", BoFormItemTypes.it_LINKED_BUTTON);
            var lblTer = form.Items.Add("lblRelPar", BoFormItemTypes.it_STATIC);
            var txtTer = form.Items.Add("txtRelpar", BoFormItemTypes.it_EDIT);
            lblTer.Left = refLbl.Left;
            lblTer.Top = refLbl.Top + 15;
            lblTer.Width = refLbl.Width;
            ((StaticText)lblTer.Specific).Caption = "Tercero Relacionado";

            txtTer.Left = refTxt.Left;
            txtTer.Top = refTxt.Top + 15;
            txtTer.Width = refTxt.Width;
            ((EditText)txtTer.Specific).DataBind.SetBound(true, form.DataSources.DBDataSources.Item(0).TableName, "U_HCO_RELPAR");
            ((EditText)txtTer.Specific).ChooseFromListUID = "CF_TERP";
            ((EditText)txtTer.Specific).ChooseFromListAlias = "Code";

            linkTer.Top = txtTer.Top;
            linkTer.Left = txtTer.Left - 20;
            linkTer.LinkTo = "txtRelpar";
            ((LinkedButton)linkTer.Specific).LinkedObject = BoLinkedObject.lf_UserDefinedObject;
            ((LinkedButton)linkTer.Specific).LinkedObjectType = "HCO_FRP1100";
        }

        public static void LoadRelatedPartiesForm(Settings.RelatedParties type, bool state = true)
        {
            try
            {
                if (type == Settings.RelatedParties.RELATED_PARTIES_MUNICIPALITY)
                {
                    var form = T1.B1.MainObject.Instance.B1Application.Forms.ActiveForm;
                    InitMunicipality(form);
                    return;
                }

                FormCreationParams objParams = (FormCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
                objParams.XmlData = GetXmlUDO(type);
                objParams.FormType = GetTypeUDO(type);
                objParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);

                Form objForm = MainObject.Instance.B1Application.Forms.AddEx(objParams);

                if (type == Settings.RelatedParties.RELATED_PARTIES_CONFIGURATION)
                {
                    InitConfiguration(objForm);
                }
                else if (type == Settings.RelatedParties.RELATED_PARTIES)
                {
                    InitRelatedParties(objForm);
                }
                else if (type == Settings.RelatedParties.RELATED_PARTIES_MOVEMENT)
                {
                    InitRelatedPartiesReport(objForm);
                }
                else if (type == Settings.RelatedParties.RELATED_PARTIES_MOVEMENT_DETAILS)
                {
                    InitThirdMovementDetails(objForm);
                }
                else if (type == Settings.RelatedParties.RELATED_PARTIES_DUMMIES)
                {
                    objForm.Left = -232;
                    objForm.Top = -231;
                    //objForm.VisibleEx = false;
                }

                objForm.VisibleEx = true;

                if (type == Settings.RelatedParties.RELATED_PARTIES_CONFIGURATION)
                {
                    MainObject.Instance.B1Application.Menus.Item("1290").Enabled = true;
                    MainObject.Instance.B1Application.Menus.Item("1290").Activate();

                    MainObject.Instance.B1Application.Forms.ActiveForm.EnableMenu("1290", false);
                    MainObject.Instance.B1Application.Forms.ActiveForm.EnableMenu("1288", false);
                    MainObject.Instance.B1Application.Forms.ActiveForm.EnableMenu("1289", false);
                    MainObject.Instance.B1Application.Forms.ActiveForm.EnableMenu("1291", false);
                }
            }
            catch (Exception er)
            {
                _Logger.Error("(LoadRelatedParties)", er);
            }
        }
        private static void InitMunicipality(Form form)
        {
            Base.UIOperations.FormsOperations.AddChooseFromList(form, "HCO_FRP0004", "cflDpto", false);
            var matrix = (Matrix)form.Items.Item("3").Specific;
            matrix.Columns.Item("U_Departamento").ChooseFromListUID = "cflDpto";
            matrix.Columns.Item("U_Departamento").ChooseFromListAlias = "Code";
        }
        private static void InitRelatedParties(Form form)
        {
            var cmboCodDpto = (ComboBox)form.Items.Item("Item_67").Specific;
            var queryCodDpto = Queries.Instance.Queries().Get("GetDepartamentType");
            var recordSet = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recordSet.DoQuery(queryCodDpto);
            for (int i = 0; i < recordSet.RecordCount; i++)
            {
                cmboCodDpto.ValidValues.Add(recordSet.Fields.Item("Code").Value.ToString(), recordSet.Fields.Item("Name").Value.ToString());
                recordSet.MoveNext();
            }
            // ((CheckBox)form.Items.Item("Item_10").Specific).Checked = true;
            B1.Base.UIOperations.FormsOperations.SetChooseFromList(form, "CFL_OCRD", "CardType", BoConditionOperation.co_NOT_EQUAL, "L");
        }

        private static void InitRelatedPartiesReport(Form form)
        {
            for (int i = 1; i <= 5; i++)
            {
                var conds = form.ChooseFromLists.Item($"CFL_Dim{i}").GetConditions();
                var cond = conds.Add();
                cond.Alias = "DimCode";
                cond.Operation = BoConditionOperation.co_EQUAL;
                cond.CondVal = i.ToString(); ;
                form.ChooseFromLists.Item($"CFL_Dim{i}").SetConditions(conds);
            }

            var condsAcct = form.ChooseFromLists.Item("CFL_CtaD").GetConditions();
            var condAcct = condsAcct.Add();
            condAcct.BracketOpenNum = 2;
            condAcct.Alias = "Frozen";
            condAcct.Operation = BoConditionOperation.co_EQUAL;
            condAcct.CondVal = "N";
            condAcct.BracketCloseNum = 1;
            condAcct.Relationship = BoConditionRelationship.cr_AND;
            condAcct = condsAcct.Add();
            condAcct.BracketOpenNum = 1;
            condAcct.Alias = "Postable";
            condAcct.Operation = BoConditionOperation.co_EQUAL;
            condAcct.CondVal = "Y";
            condAcct.BracketCloseNum = 2;

            form.ChooseFromLists.Item("CFL_CtaD").SetConditions(condsAcct);

            condsAcct = form.ChooseFromLists.Item("CFL_CtaH").GetConditions();
            condAcct = condsAcct.Add();
            condAcct.BracketOpenNum = 2;
            condAcct.Alias = "Frozen";
            condAcct.Operation = BoConditionOperation.co_EQUAL;
            condAcct.CondVal = "N";
            condAcct.BracketCloseNum = 1;
            condAcct.Relationship = BoConditionRelationship.cr_AND;
            condAcct = condsAcct.Add();
            condAcct.BracketOpenNum = 1;
            condAcct.Alias = "Postable";
            condAcct.Operation = BoConditionOperation.co_EQUAL;
            condAcct.CondVal = "Y";
            condAcct.BracketCloseNum = 2;
            form.ChooseFromLists.Item("CFL_CtaH").SetConditions(condsAcct);
        }

        public static void CheckLevelCondition(ItemEvent pVal)
        {
            var form = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            var level1 = form.DataSources.UserDataSources.Item("UD_Niv1").Value;
            var level2 = form.DataSources.UserDataSources.Item("UD_Niv2").Value;

            if (level1 == level2)
            {
                form.DataSources.UserDataSources.Item("UD_Niv2").Value = "";
            }
        }

        public static void SetReferenceChangesTypes(ItemEvent pVal)
        {
            var hash = CreateMD5(DateTime.Now.ToString("yyyyMMddhhmmss")).Substring(0, 15);
            var form = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            ((EditText)form.Items.Item("4").Specific).Value = hash;
        }

        public static void SetReferenceJournalTemplate(ItemEvent pVal)
        {
            var form = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            var valueTmplate = ((EditText)form.Items.Item("28").Specific).Value;
            var cmboTmplate = ((ComboBox)form.Items.Item("27").Specific).Selected;

            if (!string.IsNullOrEmpty(valueTmplate))
            {
                var thirdQuery = string.Format((cmboTmplate.Value == "2" ? Queries.Instance.Queries().Get("GetThirdContabPer") : Queries.Instance.Queries().Get("GetThirdContabTmpl")), valueTmplate);
                var matrixJournal = (Matrix)form.Items.Item("76").Specific;
                var record = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                record.DoQuery(thirdQuery);

                if (record.RecordCount > 0)
                {
                    try
                    {
                        form.Freeze(true);
                        for (int i = 1; i <= matrixJournal.RowCount; i++)
                        {
                            ((EditText)matrixJournal.GetCellSpecific("U_HCO_RELPAR", i)).Value = record.Fields.Item("U_HCO_RELPAR").Value.ToString();
                        }
                    }
                    finally
                    {
                        form.Freeze(false);
                    }
                }
            }
        }

        public static void SetReferencePeriodContab(ItemEvent pVal)
        {
            var hash = CreateMD5(DateTime.Now.ToString("yyyyMMddhhmmss")).Substring(0, 15);
            var form = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            if (form.Mode == BoFormMode.fm_ADD_MODE)
                ((EditText)form.Items.Item("19").Specific).Value = hash;
        }

        private static string CreateMD5(string input)
        {
            byte[] valueBytes = new byte[input.Length]; // <-- don't multiply by 2!

            var encoder = System.Text.Encoding.UTF8.GetEncoder(); // <-- UTF8 here
            encoder.GetBytes(input.ToCharArray(), 0, input.Length, valueBytes, 0, true);

            System.Security.Cryptography.MD5 md5 = new System.Security.Cryptography.MD5CryptoServiceProvider();
            byte[] hashBytes = md5.ComputeHash(valueBytes);

            var stringBuilder = new System.Text.StringBuilder();

            for (int i = 0; i < hashBytes.Length; i++)
            {
                stringBuilder.Append(hashBytes[i].ToString("x2"));
            }

            return stringBuilder.ToString();
        } 

        public static bool ValidateFieldsPeriodTempl(ItemEvent pVal)
        {
            var form = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            var third = form.DataSources.DBDataSources.Item("OTRT").GetValue("U_HCO_RELPAR", 0);
            if (form.Mode == BoFormMode.fm_ADD_MODE || form.Mode == BoFormMode.fm_UPDATE_MODE)
            {
                if (third.Equals(String.Empty))
                {
                    MainObject.Instance.B1Application.SetStatusBarMessage("No puede dejar el campo de tercero vacio");
                    return false;
                }
            }

            return true;
        }

        public static bool ValidateFieldsPeriodCont(ItemEvent pVal)
        {
            var form = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            var sel = ((ComboBox)form.Items.Item("21").Specific).Selected;
            var third = form.DataSources.DBDataSources.Item("ORCR").GetValue("U_HCO_RELPAR", 0);

            if (form.Mode == BoFormMode.fm_ADD_MODE || form.Mode == BoFormMode.fm_UPDATE_MODE)
            {
                if (third.Equals(String.Empty))
                {
                    MainObject.Instance.B1Application.SetStatusBarMessage("No puede dejar el campo de tercero vacio");
                    return false;
                }

                if (sel == null)
                {
                    MainObject.Instance.B1Application.SetStatusBarMessage("No puede dejar el campo de tipo de codigo de transaccion vacio");
                    return false;
                }
                else
                {
                    if (sel.Value == "")
                    {
                        MainObject.Instance.B1Application.SetStatusBarMessage("No puede dejar el campo de tipo de codigo de transaccion vacio");
                        return false;
                    }

                    if (sel.Value != "CP")
                    {
                        MainObject.Instance.B1Application.SetStatusBarMessage("Tiene que seleccionar el codigo de transaccion del tipo \"CP\"");
                        return false;
                    }
                }
            }

            return true;
        }

        public static bool ValidateFieldsChangesTypes(ItemEvent pVal)
        {
            var opt = true;
            var form = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            var sel = ((ComboBox)form.Items.Item("11").Specific).Selected;

            if (form.Mode == BoFormMode.fm_ADD_MODE)
            {
                var areaVal = form.DataSources.UserDataSources.Item("UD_MetVat").Value;
                if (string.IsNullOrEmpty(areaVal))
                {
                    MainObject.Instance.B1Application.SetStatusBarMessage("Debe seleccionar el area de valorizacion");
                    return false;
                }

                var autAnul = ((CheckBox)form.Items.Item("27").Specific);
                if (!autAnul.Checked)
                {
                    MainObject.Instance.B1Application.SetStatusBarMessage("Debe seleccionar las anulaciones Automaticas");
                    return false;
                }
                else
                {
                    var dateAnul = ((EditText)form.Items.Item("28").Specific).Value;
                    if (string.IsNullOrEmpty(dateAnul))
                    {
                        MainObject.Instance.B1Application.SetStatusBarMessage("Debe indicar una fecha de anulacion");
                        return false;
                    }
                }

                if (sel == null)
                {
                    opt = false;
                }
                else
                {
                    if (pVal.FormTypeEx == "369")
                    {
                        if (sel.Value == "")
                            opt = false;
                        else if (sel.Value != "HDCA")
                            opt = false;
                    }
                    else if (pVal.FormTypeEx == "371")
                    {
                        if (sel.Value == "")
                            opt = false;
                        else if (sel.Value != "DCO")
                            opt = false;
                    }
                }

                if (!opt)
                {
                    var cod = pVal.FormTypeEx == "369" ? "HDCA" : "DCO";
                    MainObject.Instance.B1Application.SetStatusBarMessage($"Debe seleccionar el codigo de transaccion {cod}");
                }
            }

            return opt;
        }

        public static bool ValidateFieldsMovement(ItemEvent pVal)
        {
            var form = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            var level1 = form.DataSources.UserDataSources.Item("UD_Niv1").Value;
            var level2 = form.DataSources.UserDataSources.Item("UD_Niv2").Value;

            if (level1 == string.Empty || level2 == string.Empty)
            {
                MainObject.Instance.B1Application.SetStatusBarMessage("Tiene que seleccionar los niveles del reporte");
                return false;
            }

            return true;
        }

        public static void OpenCrystalReport(string formuid)
        {
            MainObject.Instance.B1Application.SetStatusBarMessage("Generando movimiento de terceros, por favor espere.", BoMessageTime.bmt_Medium, false);
            try
            {
                var form = MainObject.Instance.B1Application.Forms.Item(formuid);
                var br = (SAPbobsCOM.SBObob)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoBridge);
                var fDesde = form.DataSources.UserDataSources.Item("UD_FDes").Value == "" ? "" : DateTime.Parse(br.Format_StringToDate(form.DataSources.UserDataSources.Item("UD_FDes").Value).Fields.Item(0).Value.ToString()).ToString("yyyy-MM-dd");
                var fHasta = form.DataSources.UserDataSources.Item("UD_FHas").Value == "" ? "" : DateTime.Parse(br.Format_StringToDate(form.DataSources.UserDataSources.Item("UD_FHas").Value).Fields.Item(0).Value.ToString()).ToString("yyyy-MM-dd");
                var queryDelete = $"DELETE FROM \"@HCO_CRMRP0010\" WHERE \"U_CodUsr\"= '{MainObject.Instance.B1Company.UserSignature}'";
                var insertData = "INSERT INTO \"@HCO_CRMRP0010\" (\"Code\", \"U_CodUsr\", \"U_FechaDesde\", \"U_FechaHasta\", \"U_SNDesde\", \"U_SNHasta\", \"U_CtaDesde\", \"U_CtaHasta\", \"U_NITDesde\", \"U_NITHasta\", \"U_Proyecto\", \"U_DIM1\", \"U_DIM2\", \"U_DIM3\", \"U_DIM4\", \"U_DIM5\", \"U_NIVEL1\", \"U_NIVEL2\") VALUES((CASE WHEN (SELECT MAX(\"Code\")+1 FROM \"@HCO_CRMRP0010\") IS NULL THEN 1 ELSE (SELECT MAX(\"Code\")+1 FROM \"@HCO_CRMRP0010\") END) , '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}')";
                insertData = string.Format(insertData, MainObject.Instance.B1Company.UserSignature, fDesde, fHasta, form.DataSources.UserDataSources.Item("UD_SocD").Value, form.DataSources.UserDataSources.Item("UD_SocH").Value,
                                                   form.DataSources.UserDataSources.Item("UD_CtaD").Value, form.DataSources.UserDataSources.Item("UD_CtaH").Value, form.DataSources.UserDataSources.Item("UD_NitD").Value, form.DataSources.UserDataSources.Item("UD_NitH").Value, form.DataSources.UserDataSources.Item("UD_Pro").Value,
                                                   form.DataSources.UserDataSources.Item("UD_Dim1").Value, form.DataSources.UserDataSources.Item("UD_Dim2").Value, form.DataSources.UserDataSources.Item("UD_Dim3").Value, form.DataSources.UserDataSources.Item("UD_Dim4").Value, form.DataSources.UserDataSources.Item("UD_Dim5").Value,
                                                   form.DataSources.UserDataSources.Item("UD_Niv1").Value, form.DataSources.UserDataSources.Item("UD_Niv2").Value);

                var valuesDB = GetDBValues();
                var recordSet = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                recordSet.DoQuery(queryDelete);
                recordSet.DoQuery(insertData);

                Form oForm = MainObject.Instance.B1Application.Forms.Item(formuid);
                MainObject.Instance.B1Application.Menus.Item("4873").Activate();
                var oForm2 = MainObject.Instance.B1Application.Forms.ActiveForm;
                    oForm2.Visible = false;
                ((EditText)oForm2.Items.Item("410000004").Specific).Value = Directory.GetCurrentDirectory() + "\\Report\\informe_movimiento_tercero_hana.rpt";
                oForm2.Items.Item("410000001").Click();
                var oForm3 = MainObject.Instance.B1Application.Forms.ActiveForm;
                    oForm3.Visible = false;
                ((EditText)oForm3.Items.Item("1000003").Specific).Value = MainObject.Instance.B1Company.UserSignature.ToString();               
                oForm3.Items.Item("1").Click();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Exception Crystal => " + ex.Message);
            }
            finally
            {
                MainObject.Instance.B1Application.SetStatusBarMessage("Generacion de reporte finalizado", BoMessageTime.bmt_Medium, false);
            }
        }

        private static Tuple<string, string> GetDBValues()
        {
            var queryCodDpto = Queries.Instance.Queries().Get("GetDataBaseValues");
            var recordSet = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recordSet.DoQuery(queryCodDpto);
            recordSet.MoveFirst();

            var tuple = new Tuple<string, string>(recordSet.Fields.Item("U_UserNameDB").Value.ToString(), recordSet.Fields.Item("U_PassWordDB").Value.ToString());
            return tuple;
        }

        private static void InitConfiguration(Form form)
        {
            var queryCstmr = Queries.Instance.Queries().Get("GetCustomerSeries");
            var querySupl = Queries.Instance.Queries().Get("GetSupplierSeries");
            var cmboCstmr = (ComboBox)form.Items.Item("Item_5").Specific;
            var cmboSupl = (ComboBox)form.Items.Item("Item_9").Specific;

            var recordSet = (Recordset)T1.B1.MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            recordSet.DoQuery(queryCstmr);
            for (int i = 0; i < recordSet.RecordCount; i++)
            {
                cmboCstmr.ValidValues.Add(recordSet.Fields.Item("Series").Value.ToString(), recordSet.Fields.Item("SeriesName").Value.ToString());
                recordSet.MoveNext();
            }

            recordSet.DoQuery(querySupl);
            for (int i = 0; i < recordSet.RecordCount; i++)
            {
                cmboSupl.ValidValues.Add(recordSet.Fields.Item("Series").Value.ToString(), recordSet.Fields.Item("SeriesName").Value.ToString());
                recordSet.MoveNext();
            }
        }
        public static string GetXmlUDO(Settings.RelatedParties type)
        {
            XmlDocument oXML = new XmlDocument();
            switch (type)
            {
                case Settings.RelatedParties.RELATED_PARTIES:
                    oXML.Load(AppDomain.CurrentDomain.BaseDirectory + "\\Forms\\HCO_Terceros_Relacionados.srf");
                    break;
                case Settings.RelatedParties.RELATED_PARTIES_CONFIGURATION:
                    oXML.Load(AppDomain.CurrentDomain.BaseDirectory + "\\Forms\\HCO_Terceros_Configuracion.srf");
                    break;
                case Settings.RelatedParties.RELATED_PARTIES_TYPES:
                    oXML.Load(AppDomain.CurrentDomain.BaseDirectory + "\\Forms\\HCO_Terceros_Tipos_Terceros.srf");
                    break;
                case Settings.RelatedParties.RELATED_PARTIES_DEPARTMENT:
                    oXML.Load(AppDomain.CurrentDomain.BaseDirectory + "\\Forms\\HCO_Terceros_Departamentos.srf");
                    break;
                case Settings.RelatedParties.RELATED_PARTIES_MOVEMENT:
                    oXML.Load(AppDomain.CurrentDomain.BaseDirectory + "\\Forms\\HCO_Terceros_Movimientos.srf");
                    break;
                case Settings.RelatedParties.RELATED_PARTIES_CREATION_WIZARD:
                    oXML.Load(AppDomain.CurrentDomain.BaseDirectory + "\\Forms\\HCO_CargaTerceros.srf");
                    break;
                default:
                    return string.Empty;
            }
            return oXML.InnerXml;
        }
        public static string GetTypeUDO(Settings.RelatedParties type)
        {
            switch (type)
            {
                case Settings.RelatedParties.RELATED_PARTIES:
                    return "HCO_FRP1100";
                case Settings.RelatedParties.RELATED_PARTIES_CONFIGURATION:
                    return "HCO_FRP0001";
                case Settings.RelatedParties.RELATED_PARTIES_TYPES:
                    return "HCO_T1RPA300UDO";
                case Settings.RelatedParties.RELATED_PARTIES_DEPARTMENT:
                    return "HCO_T1RPA400UDO";
                case Settings.RelatedParties.RELATED_PARTIES_MOVEMENT:
                    return "HCO_T1RPA500UDO";
                case Settings.RelatedParties.RELATED_PARTIES_CREATION_WIZARD:
                    return "HCO_FRP2100";
                default:
                    return string.Empty;
            }
        }
        static public void DeleteBP()
        {
            var form = MainObject.Instance.B1Application.Forms.ActiveForm;
            var bp = (BusinessPartners)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
            var code = form.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0);
            var type = form.DataSources.DBDataSources.Item(0).GetValue("CardType", 0);
            var queryCount = string.Format(Queries.Instance.Queries().Get("CheckBPThirdCant"), code, type);
            var record = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            record.DoQuery(queryCount);
            record.MoveFirst();



            if (record.Fields.Item("Cant").Value.ToString().Equals("0") || record.Fields.Item("Cant").Value.ToString().Equals(string.Empty))
            {
                if (bp.GetByKey(code))
                {
                    if (bp.Remove() != 0)
                    {
                        MainObject.Instance.B1Application.SetStatusBarMessage(MainObject.Instance.B1Company.GetLastErrorDescription());
                    }
                    else
                    {
                        MainObject.Instance.B1Application.SetStatusBarMessage("Socio de negocios eliminado exitosamente.", BoMessageTime.bmt_Medium, false);
                        MainObject.Instance.B1Application.Menus.Item("1281").Activate();
                    }
                }
            }
            else
            {
                MainObject.Instance.B1Application.SetStatusBarMessage("No se puede eliminar el socio de negocio si tiene terceros relacionados.");
            }
        }
        static public void LoadCrystalReport(TYPE_CRYSTAL type)
        {
            MainObject.Instance.B1Application.ActivateMenuItem("4873");
            var form = MainObject.Instance.B1Application.Forms.ActiveForm;
            ((EditText)form.Items.Item("410000004").Specific).Value = GetTypeReport(type);
            form.Items.Item("410000001").Click();
        }

        static private string GetTypeReport(TYPE_CRYSTAL type)
        {
            switch (type)
            {
                case TYPE_CRYSTAL.AUXILIAR:
                    return AppDomain.CurrentDomain.BaseDirectory + "\\Report\\Auxiliar de cuenta.rpt";
                case TYPE_CRYSTAL.BALANCE:
                    return AppDomain.CurrentDomain.BaseDirectory + "\\Report\\Balance de prueba.rpt";
                case TYPE_CRYSTAL.BALNCE_TEST_RP:
                    return AppDomain.CurrentDomain.BaseDirectory + "\\Report\\Balance de prueba por tercero.rpt";
                case TYPE_CRYSTAL.DIARIO:
                    return AppDomain.CurrentDomain.BaseDirectory + "\\Report\\Movimiento Diario.rpt";

                case TYPE_CRYSTAL.ESFA:
                    return AppDomain.CurrentDomain.BaseDirectory + "\\Report\\ESF(BALANCE).rpt";
                case TYPE_CRYSTAL.ERI:
                    return AppDomain.CurrentDomain.BaseDirectory + "\\Report\\ERI (Perdidas y ganancias).rpt";
                case TYPE_CRYSTAL.TERCERO:
                    return AppDomain.CurrentDomain.BaseDirectory + "\\Report\\Movimiento terceros.rpt";

                case TYPE_CRYSTAL.IVASALE_COD:
                    return AppDomain.CurrentDomain.BaseDirectory + "\\Report\\IVA_VentasPorCodigo.rpt";
                case TYPE_CRYSTAL.IVAPRUCH_COD:
                    return AppDomain.CurrentDomain.BaseDirectory + "\\Report\\IVA_ComprasPorCodigo.rpt";

                case TYPE_CRYSTAL.RETPURCH_COD:
                    return AppDomain.CurrentDomain.BaseDirectory + "\\Report\\RetencionesComprasPorCodigo.rpt";
                case TYPE_CRYSTAL.RETSALE_COD:
                    return AppDomain.CurrentDomain.BaseDirectory + "\\Report\\RetencionesVentasPorCodigo.rpt";
                case TYPE_CRYSTAL.RETPRUCH_CARD:
                    return AppDomain.CurrentDomain.BaseDirectory + "\\Report\\RetencionesComprasPorProveedor.rpt";
                case TYPE_CRYSTAL.RETSALE_CARD:
                    return AppDomain.CurrentDomain.BaseDirectory + "\\Report\\RetencionesVentasPorCliente.rpt";

                case TYPE_CRYSTAL.CERT_RET:
                    return AppDomain.CurrentDomain.BaseDirectory + "\\Report\\CertificadoRetencionCompras.rpt";
                case TYPE_CRYSTAL.CERT_RET_IVA:
                    return AppDomain.CurrentDomain.BaseDirectory + "\\Report\\CertificadoRetencionComprasIVA.rpt";
                default:
                    return string.Empty;
            }
        }

        static public void DeleteThirdParty()
        {

            if (MainObject.Instance.B1Application.MessageBox("¿Está seguro que desea eliminar definitivamente el socio de negocio? Esta acción no se puede deshacer", 1, "Si", "No") == 1)
            {
                MainObject.Instance.B1Company.StartTransaction();

                var form = MainObject.Instance.B1Application.Forms.ActiveForm;
                var code = form.DataSources.DBDataSources.Item(0).GetValue("Code", 0);
                var queryCount = string.Format(Queries.Instance.Queries().Get("GetBPCount"), code);
                var record = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                record.DoQuery(queryCount);
                record.MoveFirst();

                try
                {
                    if (record.Fields.Item("Cant").Value.ToString().Equals("0") || record.Fields.Item("Cant").Value.ToString().Equals(string.Empty))
                    {
                        DeleteThird(form, code);
                        MainObject.Instance.B1Company.EndTransaction(BoWfTransOpt.wf_Commit);
                    }
                    else
                    {
                        if (!DeleteThirdAssociated(code))
                        {
                            DeleteThird(form, code);
                            MainObject.Instance.B1Application.SetStatusBarMessage("No puede eliminar el tercero si tiene socios de negocios relacionados o verifique el siguiente error => " + MainObject.Instance.B1Company.GetLastErrorDescription());
                            MainObject.Instance.B1Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                        }
                        else
                        {
                            if (DeleteThird(form, code))
                            {
                                MainObject.Instance.B1Application.SetStatusBarMessage("Tercero eliminado exitosamente", BoMessageTime.bmt_Medium, false);
                                MainObject.Instance.B1Company.EndTransaction(BoWfTransOpt.wf_Commit);
                            }
                            else
                                MainObject.Instance.B1Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                        }
                    }
                }
                catch
                {
                    MainObject.Instance.B1Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
            }
        }
        static public void DeleteThirdParty(Form oForm, string codeBP)
        {
            MainObject.Instance.B1Company.StartTransaction();

            try
            {
                if (!DeleteBPAssociated(codeBP))
                {
                    MainObject.Instance.B1Application.SetStatusBarMessage("No se puede eliminar el tercero si tiene socios de negocios relacionados o verifique el siguiente error => " + MainObject.Instance.B1Company.GetLastErrorDescription());
                    MainObject.Instance.B1Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
                else
                {
                    MainObject.Instance.B1Company.EndTransaction(BoWfTransOpt.wf_Commit);
                }
            }
            catch
            {
                MainObject.Instance.B1Company.EndTransaction(BoWfTransOpt.wf_RollBack);
            }
        }
        static public void RelatedPartiedMatrixOperationUDO(string Action, Form oForm)
        {
            try
            {
                var objMatrix = (Matrix)oForm.Items.Item("Item_52").Specific;

                switch (Action)
                {
                    case "Add":
                        oForm.DataSources.DBDataSources.Item(1).InsertRecord(oForm.DataSources.DBDataSources.Item(1).Size);
                        oForm.DataSources.DBDataSources.Item(1).Offset = oForm.DataSources.DBDataSources.Item(1).Size - 1;
                        objMatrix.AddRow(1);

                        for (int i = 1; i <= objMatrix.RowCount; i++)
                            objMatrix.SetCellWithoutValidation(objMatrix.RowCount, "#", i.ToString());

                        break;
                    case "Delete":

                        for (int i = objMatrix.RowCount; i >= 1; i--)
                            if (objMatrix.IsRowSelected(i))
                                objMatrix.DeleteRow(i);

                        var numerationUID = objMatrix.Columns.Item(0).UniqueID;
                        for (int i = 1; i <= objMatrix.RowCount; i++)
                            ((EditText)objMatrix.GetCellSpecific(numerationUID, i)).Value = i.ToString();

                        break;
                }

                if (oForm.Mode == BoFormMode.fm_OK_MODE)
                    oForm.Mode = BoFormMode.fm_UPDATE_MODE;
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

        private static void InitThirdMovementDetails(Form form)
        {
            var grid = (Grid)form.Items.Item("Item_0").Specific;
            var formCurrent = MainObject.Instance.B1Application.Forms.ActiveForm;
            form.DataSources.UserDataSources.Item("UD_FecD").Value = formCurrent.DataSources.UserDataSources.Item("UD_FDes").Value;
            form.DataSources.UserDataSources.Item("UD_FecH").Value = formCurrent.DataSources.UserDataSources.Item("UD_FHas").Value;
            form.DataSources.UserDataSources.Item("UD_SocD").Value = formCurrent.DataSources.UserDataSources.Item("UD_SocD").Value;
            form.DataSources.UserDataSources.Item("UD_SocH").Value = formCurrent.DataSources.UserDataSources.Item("UD_SocH").Value;
            form.DataSources.UserDataSources.Item("UD_NitD").Value = formCurrent.DataSources.UserDataSources.Item("UD_NitD").Value;
            form.DataSources.UserDataSources.Item("UD_NitH").Value = formCurrent.DataSources.UserDataSources.Item("UD_NitH").Value;
            form.DataSources.UserDataSources.Item("UD_Proy").Value = formCurrent.DataSources.UserDataSources.Item("UD_Pro").Value;
            form.DataSources.UserDataSources.Item("UD_CtaD").Value = formCurrent.DataSources.UserDataSources.Item("UD_CtaD").Value;
            form.DataSources.UserDataSources.Item("UD_CtaH").Value = formCurrent.DataSources.UserDataSources.Item("UD_CtaH").Value;
            form.DataSources.UserDataSources.Item("UD_Niv1").Value = formCurrent.DataSources.UserDataSources.Item("UD_Niv1").Value;
            form.DataSources.UserDataSources.Item("UD_Niv2").Value = formCurrent.DataSources.UserDataSources.Item("UD_Niv2").Value;
            form.DataSources.UserDataSources.Item("UD_Dim1").Value = formCurrent.DataSources.UserDataSources.Item("UD_Dim1").Value;
            form.DataSources.UserDataSources.Item("UD_Dim2").Value = formCurrent.DataSources.UserDataSources.Item("UD_Dim2").Value;
            form.DataSources.UserDataSources.Item("UD_Dim3").Value = formCurrent.DataSources.UserDataSources.Item("UD_Dim3").Value;
            form.DataSources.UserDataSources.Item("UD_Dim4").Value = formCurrent.DataSources.UserDataSources.Item("UD_Dim4").Value;
            form.DataSources.UserDataSources.Item("UD_Dim5").Value = formCurrent.DataSources.UserDataSources.Item("UD_Dim5").Value;

            var query = string.Format(Queries.Instance.Queries().Get("GetQueryThirdMovement2"), "");
            grid.DataTable.ExecuteQuery(query);
            grid.Item.Enabled = false;
        }
        static public void CheckLinesUDO(BusinessObjectInfo objBusinessObjectInfo)
        {
            var form = MainObject.Instance.B1Application.Forms.Item(objBusinessObjectInfo.FormUID);
            var dbLines = form.DataSources.DBDataSources.Item("@HCO_RP1101");
            for (int i = dbLines.Size - 1; i >= 0; i--)
            {
                if (dbLines.GetValue("U_CardCode", i).Equals(string.Empty))
                    dbLines.RemoveRecord(i);
            }
        }
        static public void FrozenForRelParty(SAPbouiCOM.BusinessObjectInfo objBusinessObjectInfo)
        {
            try
            {
                var form = MainObject.Instance.B1Application.Forms.Item(objBusinessObjectInfo.FormUID);
                var checkActive = form.DataSources.DBDataSources.Item(0).GetValue("U_Active", 0).Trim().Equals("Y") ? true : false;

                SetBPStatusCheck(form, checkActive);
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);
                MainObject.Instance.B1Application.SetStatusBarMessage("HCO:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("HCO:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }

        }
        static public void SetBPStatusCheck(Form form, bool state)
        {
            try
            {
                var sn = (BusinessPartners)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                var matrix = (Matrix)form.Items.Item("Item_52").Specific;
                for (int i = 1; i <= matrix.VisualRowCount; i++)
                {
                    var codigo = ((EditText)matrix.GetCellSpecific("Col_0", i)).Value;
                    if (!string.IsNullOrEmpty(codigo))
                    {
                        if (sn.GetByKey(codigo))
                        {
                            sn.Frozen = !state ? BoYesNoEnum.tYES : BoYesNoEnum.tNO;
                            if (!state)
                                sn.Valid = BoYesNoEnum.tNO;
                            else
                                sn.Valid = BoYesNoEnum.tYES;

                            var resp = sn.Update();
                            var msg = MainObject.Instance.B1Company.GetLastErrorDescription();
                        }
                    }
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);
                MainObject.Instance.B1Application.SetStatusBarMessage("HCO:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("HCO:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }
        static public void SetVerificationDigit(string formUID)
        {
            try
            {
                var form = MainObject.Instance.B1Application.Forms.Item(formUID);
                var doctype = form.DataSources.DBDataSources.Item(0).GetValue("U_DocTypeID", 0);
                var nit = form.DataSources.DBDataSources.Item(0).GetValue("U_LicTradNum", 0);

                if (nit.Equals(string.Empty))
                    form.DataSources.DBDataSources.Item(0).SetValue("U_AuthDig", 0, string.Empty);

                if (doctype.Equals("31") || doctype.Equals("13"))
                {
                    var digit = GetVerificationDigit(nit);
                    if (digit != -1)
                        form.DataSources.DBDataSources.Item(0).SetValue("U_AuthDig", 0, digit.ToString());
                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);
                MainObject.Instance.B1Application.SetStatusBarMessage("HCO:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("HCO:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }

        }
        static private int GetVerificationDigit(string nit)
        {
            var digitos = new byte[15];
            for (int i = 0; i < 15; i++)
            {
                digitos[i] = 0;
            }

            int cont = 0;
            int init = 15 - nit.Length;
            for (int i = init; i < 15; i++)
            {
                if (!char.IsDigit(nit[cont]))
                    return -1;
                digitos[i] = byte.Parse(nit[cont].ToString());
                cont++;
            }


            var v =
                (3 * digitos[14]) +
                (7 * digitos[13]) +
                (13 * digitos[12]) +
                (17 * digitos[11]) +
                (19 * digitos[10]) +
                (23 * digitos[9]) +
                (29 * digitos[8]) +
                (37 * digitos[7]) +
                (41 * digitos[6]) +
                (43 * digitos[5]) +
                (47 * digitos[4]) +
                (53 * digitos[3]) +
                (59 * digitos[2]) +
                (67 * digitos[1]) +
                (71 * digitos[0]);
            v = v % 11;
            if (v > 1)
                v = 11 - v;

            return v;
        }
        static public int GetNext_BPSecuence(TYPE_BP type)
        {
            var form = MainObject.Instance.B1Application.Forms.ActiveForm;
            var code = form.DataSources.DBDataSources.Item(0).GetValue("U_LicTradNum", 0) + 
                ((!string.IsNullOrEmpty(form.DataSources.DBDataSources.Item(0).GetValue("U_AuthDig", 0))) ? "-" + form.DataSources.DBDataSources.Item(0).GetValue("U_AuthDig", 0) : "");
            var queryCheck = string.Format(Queries.Instance.Queries().Get("CheckBPCant"), code, type == TYPE_BP.CUSTOMER ? "C" : "S");
            var record = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            record.DoQuery(queryCheck);
            record.MoveFirst();

            return Int32.Parse(record.Fields.Item("Sec").Value.ToString());

        }
        static public void DeleteRowBP()
        {
            MainObject.Instance.B1Company.StartTransaction();

            try
            {
                var form = MainObject.Instance.B1Application.Forms.ActiveForm;
                var matrix = (Matrix)form.Items.Item("Item_52").Specific;
                var bp = (BusinessPartners)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                var code = form.DataSources.DBDataSources.Item(0).GetValue("Code", 0);

                for (int i = 1; i <= matrix.VisualRowCount; i++)
                {
                    if (matrix.IsRowSelected(i))
                    {
                        var bpCode = ((EditText)matrix.GetCellSpecific("Col_0", i)).Value;
                        if (bp.GetByKey(bpCode))
                        {
                            if (bp.Remove() == 0)
                            {
                                deleteDataUDO(code, bpCode);
                                MainObject.Instance.B1Company.EndTransaction(BoWfTransOpt.wf_Commit);
                                return;
                            }
                        }
                    }
                }
                
                MainObject.Instance.B1Company.EndTransaction(BoWfTransOpt.wf_Commit);
                
            }
            catch (COMException comEx)
            {
                MainObject.Instance.B1Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);
                MainObject.Instance.B1Application.SetStatusBarMessage("HCO:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            catch (Exception er)
            {
                MainObject.Instance.B1Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("HCO:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }
        static public void createBP(TYPE_BP type, int secuence)
        {
            RelatedParties.Instance.GetRelPartyConfiguration();
            var form = MainObject.Instance.B1Application.Forms.ActiveForm;
            var code = form.DataSources.DBDataSources.Item(0).GetValue("Code", 0);
            var queryConfig = Queries.Instance.Queries().Get("GetConfig");
            var bp = (BusinessPartners)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
            var record = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            int series = 0;
            int consecC = 0;
            int consecV = 0;
            string cardCode = "";
            try
            {
                record.DoQuery(queryConfig);
                record.MoveFirst();
                consecC = record.Fields.Item("U_ConsecC").Value.ToString().Equals(string.Empty) ? 0 : Int32.Parse(record.Fields.Item("U_ConsecC").Value.ToString());
                consecV = record.Fields.Item("U_ConsecV").Value.ToString().Equals(string.Empty) ? 0 : Int32.Parse(record.Fields.Item("U_ConsecV").Value.ToString());
                var consec = (type == TYPE_BP.CUSTOMER) ? consecC : consecV;

                if (MainObject.Instance.B1Application.MessageBox("Desea crear el socio de negocios ?", 1, "Si", "No") == 1)
                {

                    var manBpSerie =  CacheManager.CacheManager.Instance.getFromCache("ManBPSerie");
                    var multBP =  CacheManager.CacheManager.Instance.getFromCache("MultBP");
                    var autoConse =  CacheManager.CacheManager.Instance.getFromCache("AutoConse");
                    var terPrefix =  CacheManager.CacheManager.Instance.getFromCache("TerPerfix");
                    var prefix =  CacheManager.CacheManager.Instance.getFromCache((type == TYPE_BP.CUSTOMER) ? "ClientPrefix" : "VendorPrefix");
                    if (manBpSerie.Equals("Y"))
                    {
                        if (type == TYPE_BP.CUSTOMER)
                            series = string.IsNullOrEmpty(CacheManager.CacheManager.Instance.getFromCache("CSeries").ToString()) ? 0 : int.Parse(CacheManager.CacheManager.Instance.getFromCache("CSeries"));
                        else if(type == TYPE_BP.SUPPLIER)
                            series = string.IsNullOrEmpty(CacheManager.CacheManager.Instance.getFromCache("VSeries").ToString()) ? 0 : int.Parse(CacheManager.CacheManager.Instance.getFromCache("VSeries"));

                        if (series == 0)
                            throw new Exception("Seleccione una serie en la configuración de localización");
                    }
                    else
                    {
                        if (terPrefix.Equals("Y"))
                        {
                            var queryP = Queries.Instance.Queries().Get("GetRPPrefix");
                            record.DoQuery(string.Format(queryP, code));
                            record.MoveFirst();
                            if(record.RecordCount > 0)
                            {
                                if (autoConse.Equals("Y"))
                                    consec = string.IsNullOrEmpty(record.Fields.Item("U_Consecutive").Value.ToString()) ? 1 : int.Parse(record.Fields.Item("U_Consecutive").Value.ToString());
                                prefix = record.Fields.Item("U_Prefix").Value;
                            }                  

                            if(string.IsNullOrEmpty(prefix))
                                throw new Exception("la opción de configuración 'Manejar prefijo por tipo de tercero' está marcada pero no se configuró un prefijo para el tipo de tercero");
                        }
                        if (multBP.Equals("Y"))
                        {
                            cardCode = prefix + form.DataSources.DBDataSources.Item(0).GetValue("Code", 0) + "_" + secuence;
                        }
                        else if (autoConse.Equals("Y"))
                        {
                            var numChara = CacheManager.CacheManager.Instance.getFromCache("NumChara");
                            cardCode = prefix + consec.ToString().PadLeft(numChara, '0');
                            while (bp.GetByKey(cardCode))
                            {
                                consec += 1;
                                cardCode = prefix + consec.ToString().PadLeft(numChara, '0');
                            }
                            bp = null;
                            bp = (BusinessPartners)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                        }
                        else
                        {
                            cardCode = prefix + code;
                            if (bp.GetByKey(cardCode))
                                throw new Exception("El socio de negocios ya existe, no es posible su creación");
                        }
                    }

                    bp.CardType = (type == TYPE_BP.CUSTOMER) ? BoCardTypes.cCustomer : BoCardTypes.cSupplier;
                    if (series != 0) bp.Series = series; else bp.CardCode = cardCode;
                    bp.CardName = form.DataSources.DBDataSources.Item(0).GetValue("Name", 0);
                    bp.FederalTaxID = form.DataSources.DBDataSources.Item(0).GetValue("U_LicTradNum", 0) + 
                        (!string.IsNullOrEmpty(form.DataSources.DBDataSources.Item(0).GetValue("U_AuthDig", 0)) ? ("-" + form.DataSources.DBDataSources.Item(0).GetValue("U_AuthDig", 0)) : "");
                    bp.Phone1 = form.DataSources.DBDataSources.Item(0).GetValue("U_Phone1", 0);
                    bp.Phone2 = form.DataSources.DBDataSources.Item(0).GetValue("U_Phone2", 0);
                    bp.EmailAddress = form.DataSources.DBDataSources.Item(0).GetValue("U_Email", 0);

                    bp.Addresses.AddressName = "Principal";
                    bp.Addresses.Street = form.DataSources.DBDataSources.Item(0).GetValue("U_MainAddress", 0);
                    bp.Addresses.Country = form.DataSources.DBDataSources.Item(0).GetValue("U_CountryCode", 0);
                    bp.Addresses.UserFields.Fields.Item("U_HCO_MUNI").Value = form.DataSources.DBDataSources.Item(0).GetValue("U_CountyCode", 0);
                    bp.Addresses.AddressType = BoAddressType.bo_BillTo;
                    bp.Addresses.Add();
                    bp.Addresses.AddressName = "Principal";
                    bp.Addresses.Street = form.DataSources.DBDataSources.Item(0).GetValue("U_MainAddress", 0);
                    bp.Addresses.Country = form.DataSources.DBDataSources.Item(0).GetValue("U_CountryCode", 0);
                    bp.Addresses.UserFields.Fields.Item("U_HCO_MUNI").Value = form.DataSources.DBDataSources.Item(0).GetValue("U_CountyCode", 0);
                    bp.Addresses.AddressType = BoAddressType.bo_ShipTo;

                    if (bp.Add() == 0)
                    {
                        if (autoConse.Equals("Y"))
                        {
                            if (terPrefix.Equals("Y"))
                                updateConsecRelPartyType(form.DataSources.DBDataSources.Item(0).GetValue("U_CardTypeID", 0), consec);
                            else
                                updateConsecRelPartyConfiguration(type, consec);
                        }

                        var msg = ((type == TYPE_BP.CUSTOMER) ? "Cliente" : "Proveedor") + " creado exitosamente.";

                        _Logger.Debug(msg);
                        addDataIntoUDO(type, code, MainObject.Instance.B1Company.GetNewObjectKey(), bp.CardName);

                        MainObject.Instance.B1Application.ActivateMenuItem("1304");
                        MainObject.Instance.B1Application.SetStatusBarMessage(msg, BoMessageTime.bmt_Medium, false);
                    }
                    else
                    {
                        var msg = "Error al crear el bp => " + bp.CardCode + " - " + MainObject.Instance.B1Company.GetLastErrorDescription();
                        _Logger.Debug(msg);

                        MainObject.Instance.B1Application.SetStatusBarMessage(msg);
                    }
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);
                MainObject.Instance.B1Application.SetStatusBarMessage("HCO:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("HCO:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        static private void updateConsecRelPartyConfiguration(TYPE_BP type, int concec)
        {
            try
            {
                var cs = MainObject.Instance.B1Company.GetCompanyService();
                var gs = cs.GetGeneralService("HCO_FRP0001");
                GeneralDataParams gdp = (GeneralDataParams)gs.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                gdp.SetProperty("Code", CacheManager.CacheManager.Instance.getFromCache("CodeRPConf").ToString());
                var gd = gs.GetByParams(gdp);
                gd.SetProperty((type == TYPE_BP.CUSTOMER) ? "U_ConsecC" : "U_ConsecV", concec + 1);

                gs.Update(gd);
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

        static private void updateConsecRelPartyType(string docTypeID, int concec)
        {
            try
            {
                var cs = MainObject.Instance.B1Company.GetCompanyService();
                var gs = cs.GetGeneralService("HCO_FRP0002");
                GeneralDataParams gdp = (GeneralDataParams)gs.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                gdp.SetProperty("Code", docTypeID);
                var gd = gs.GetByParams(gdp);
                gd.SetProperty("U_Consecutive", concec + 1);

                gs.Update(gd);
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
        static private bool DeleteThird(Form form, string Code)
        {
            try
            {
                if (DeleteThirdAssociated(Code))
                {
                    CompanyService cs = (CompanyService)MainObject.Instance.B1Company.GetCompanyService();
                    GeneralService gs = cs.GetGeneralService("HCO_FRP1100");
                    GeneralDataParams gdp = (GeneralDataParams)gs.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                    gdp.SetProperty("Code", Code);

                    gs.Delete(gdp);

                    MainObject.Instance.B1Application.Menus.Item("1282").Activate();
                    return true;
                }
                else
                    return false;
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error(string.Empty, comEx);
                MainObject.Instance.B1Application.SetStatusBarMessage("Error al eliminar el tercero => " + comEx);
                return false;
            }
            catch (Exception er)
            {
                _Logger.Error(string.Empty, er);
                MainObject.Instance.B1Application.SetStatusBarMessage("Error al eliminar el tercero => " + er);
                return false;
            }
        }
        private static bool DeleteBPAssociated(string cardCode)
        {
            try
            {
                var bp = (BusinessPartners)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                if (bp.GetByKey(cardCode))
                {
                    if (bp.GetByKey(cardCode))
                    {
                        if (bp.Remove() != 0)
                            return false;
                    }
                }
                return true;
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);
                return false;

            }
            catch (Exception ex)
            {
                _Logger.Error("", ex);
                MainObject.Instance.B1Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                return false;
            }
        }
        private static bool DeleteThirdAssociated(string code)
        {
            try
            {
                var queryLinesBP = string.Format(Queries.Instance.Queries().Get("GetLinesBP"), code);
                var bp = (BusinessPartners)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                var record = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                record.DoQuery(queryLinesBP);

                while (!record.EoF)
                {
                    var bpCode = (string)record.Fields.Item("BP").Value;
                    if (bp.GetByKey(bpCode))
                    {
                        if (bp.GetByKey(bpCode))
                        {
                            if (bp.Remove() != 0)
                                return false;
                        }
                    }

                    record.MoveNext();
                }

                return true;
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);
                return false;

            }
            catch (Exception ex)
            {
                _Logger.Error("", ex);
                MainObject.Instance.B1Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                return false;
            }
        }
        static private void deleteDataUDO(string Code, string CardCode)
        {
            try
            {
                var cs = MainObject.Instance.B1Company.GetCompanyService();
                var gs = cs.GetGeneralService("HCO_FRP1100");
                GeneralDataParams gdp = (GeneralDataParams)gs.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                gdp.SetProperty("Code", Code);
                var gd = gs.GetByParams(gdp);
                var gdc = (GeneralDataCollection)gd.Child("HCO_RP1101");

                for (int i = 0; i < gdc.Count; i++)
                {
                    var gdcl = gdc.Item(i);
                    if (gdcl.GetProperty("U_CardCode").ToString().Equals(CardCode))
                    {
                        gdc.Remove(i);
                    }
                }

                gs.Update(gd);
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
        static private void addDataIntoUDO(TYPE_BP type, string Code, string CardCode, string CardName)
        {
            try
            {
                var cs = MainObject.Instance.B1Company.GetCompanyService();
                var gs = cs.GetGeneralService("HCO_FRP1100");
                GeneralDataParams gdp = (GeneralDataParams)gs.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                gdp.SetProperty("Code", Code);
                var gd = gs.GetByParams(gdp);
                var gdc = gd.Child("HCO_RP1101");

                var gdcl = gdc.Add();
                gdcl.SetProperty("U_CardType", type == TYPE_BP.CUSTOMER ? "C" : "S");
                gdcl.SetProperty("U_CardCode", CardCode);
                gdcl.SetProperty("U_CardName", CardName);

                gs.Update(gd);
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
        static public void relatedPartiedMatrixOperation(EventInfoClass eventInfo, string Action)
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.Matrix objMatrix = null;
            int intTotalLines = -1;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(eventInfo.FormUID);
                objMatrix = (Matrix)objForm.Items.Item(Settings._Main.BPFormMatrixId).Specific;
                intTotalLines = objMatrix.RowCount;
                int intRow = eventInfo.Row;
                switch (Action)
                {
                    case "Add":
                        objMatrix.AddRow(1, intRow);

                        objMatrix.SetCellWithoutValidation(intRow + 1, "Col_0", "");
                        objMatrix.FlushToDataSource();

                        objMatrix.SetCellFocus(intRow + 1, 1);
                        if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }

                        break;
                    case "Delete":
                        objMatrix.DeleteRow(intRow);
                        objMatrix.FlushToDataSource();
                        if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                        break;
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


        static public bool ProcessFile(Form oForm)
        {
            int _line = 0;
            string line = string.Empty;
            string path = oForm.DataSources.UserDataSources.Item("UD_URL").Value.ToString();
            StreamReader file = new StreamReader(path, Encoding.Default);
            var delimiters = new char[] { '\t' };
            SAPbouiCOM.DataTable DT = oForm.DataSources.DataTables.Item("DT_RP");
            Matrix MT = (Matrix)oForm.Items.Item("Item_6").Specific;

            try
            {

                MT.Columns.Item("LineId").DataBind.Bind("DT_RP", "LineId");
                MT.Columns.Item("Col_0").DataBind.Bind("DT_RP", "CardCode");
                MT.Columns.Item("Col_1").DataBind.Bind("DT_RP", "CardName");
                MT.Columns.Item("Col_2").DataBind.Bind("DT_RP", "CardType");
                MT.Columns.Item("Col_3").DataBind.Bind("DT_RP", "DocType");
                MT.Columns.Item("Col_4").DataBind.Bind("DT_RP", "LicTradNum");
                MT.Columns.Item("Col_5").DataBind.Bind("DT_RP", "TaxCardType");
                MT.Columns.Item("Col_6").DataBind.Bind("DT_RP", "TaxRegim");
                MT.Columns.Item("Col_7").DataBind.Bind("DT_RP", "EconActvty");
                MT.Columns.Item("Col_8").DataBind.Bind("DT_RP", "Country");
                MT.Columns.Item("Col_9").DataBind.Bind("DT_RP", "County");
                MT.Columns.Item("Col_10").DataBind.Bind("DT_RP", "MainAddress");
                MT.Columns.Item("Col_11").DataBind.Bind("DT_RP", "Phone1");
                MT.Columns.Item("Col_12").DataBind.Bind("DT_RP", "Phone2");
                MT.Columns.Item("Col_13").DataBind.Bind("DT_RP", "Email");
                MT.Columns.Item("Col_14").DataBind.Bind("DT_RP", "ZipCode");
                MT.Columns.Item("Col_15").DataBind.Bind("DT_RP", "FirstName");
                MT.Columns.Item("Col_16").DataBind.Bind("DT_RP", "MiddleName");
                MT.Columns.Item("Col_17").DataBind.Bind("DT_RP", "LastName");
                MT.Columns.Item("Col_18").DataBind.Bind("DT_RP", "ScndSrName");
                MT.Columns.Item("Col_19").DataBind.Bind("DT_RP", "AddNames");
                MT.Columns.Item("Col_20").DataBind.Bind("DT_RP", "CardTypeC");
                MT.Columns.Item("Col_20").Visible = false;
                MT.Columns.Item("Col_21").DataBind.Bind("DT_RP", "DocTypeC");
                MT.Columns.Item("Col_21").Visible = false;
                MT.Columns.Item("Col_22").DataBind.Bind("DT_RP", "TaxCardTypeC");
                MT.Columns.Item("Col_22").Visible = false;
                MT.Columns.Item("Col_23").DataBind.Bind("DT_RP", "TaxRegimC");
                MT.Columns.Item("Col_23").Visible = false;
                MT.Columns.Item("Col_24").DataBind.Bind("DT_RP", "CountryC");
                MT.Columns.Item("Col_24").Visible = false;
                MT.Columns.Item("Col_25").DataBind.Bind("DT_RP", "CountyC");
                MT.Columns.Item("Col_25").Visible = false;
                MT.Columns.Item("Col_26").DataBind.Bind("DT_RP", "DeptC");
                MT.Columns.Item("Col_26").Visible = false;
                MT.Columns.Item("Col_27").DataBind.Bind("DT_RP", "Departamento");
                
                while ((line = file.ReadLine()) != null)
                {
                    
                    if(_line > 1)
                    {
                        DT.Rows.Add(1);
                        
                        var segments = line.Split(delimiters);
                        int _col = 1;
                        DT.SetValue(0, _line - 2, _line - 1);
                        foreach (var segment in segments)
                        {

                            if (_col == 3)
                            {
                                DT.SetValue(_col, _line - 2, GetNameValue(segment,"DT_CardType",oForm));
                                DT.SetValue("CardTypeC", _line - 2, segment);
                            } else if (_col == 4)
                            {
                                DT.SetValue(_col, _line - 2, GetNameValue(segment, "DT_DocType", oForm));
                                DT.SetValue("DocTypeC", _line - 2, segment);
                            }
                            else if (_col == 6)
                            {
                                DT.SetValue(_col, _line - 2, GetNameValue(segment, "DT_TcrdType", oForm));
                                DT.SetValue("TaxCardTypeC", _line - 2, segment);
                            }
                            else if (_col == 7)
                            {
                                DT.SetValue(_col, _line - 2, GetNameValue(segment, "DT_TaxRegim", oForm));
                                DT.SetValue("TaxRegimC", _line - 2, segment);
                            }
                            else if (_col == 9)
                            {
                                DT.SetValue(_col, _line - 2, GetNameValue(segment, "DT_Country", oForm));
                                DT.SetValue("CountryC", _line - 2, segment);
                            }
                            else if (_col == 10)
                            {
                                DT.SetValue(_col, _line - 2, GetNameValue(segment, "DT_County", oForm));
                                DT.SetValue("CountyC", _line - 2, segment);
                                string DeptCode = GetDeptCode(segment, "DT_County", oForm);                                
                                DT.SetValue("DeptC", _line - 2,DeptCode);
                                DT.SetValue("Departamento", _line - 2, GetNameValue(DeptCode, "DT_Dept", oForm));
                            }
                            else
                            {
                                DT.SetValue(_col, _line - 2, segment);
                            }
                            _col++;
                        }
                    }
                    _line++;
                }
                MT.LoadFromDataSourceEx();
                MT.AutoResizeColumns();

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

            return true;

        }

        private static string GetNameValue(string segment, string dataTable, Form oform)
        {
            SAPbouiCOM.DataTable DT = oform.DataSources.DataTables.Item(dataTable);
            string valueName = string.Empty;

            var data = B1.Base.UIOperations.FormsOperations.SapDataTableToDotNetDataTable(DT.SerializeAsXML(BoDataTableXmlSelect.dxs_All));

            DataRow[] result = data.Select("Code = '" + segment + "'");

            if (result.Length > 0)
            {
                valueName = result[0][1].ToString();
            }

            return valueName;
        }
        private static string GetDeptCode(string segment, string dataTable, Form oform)
        {
            SAPbouiCOM.DataTable DT = oform.DataSources.DataTables.Item(dataTable);
            string DeptCode = string.Empty;

            var data = B1.Base.UIOperations.FormsOperations.SapDataTableToDotNetDataTable(DT.SerializeAsXML(BoDataTableXmlSelect.dxs_All));

            DataRow[] result = data.Select("Code = '" + segment + "'");

            if (result.Length > 0)
            {
                DeptCode = result[0][2].ToString();
            }

            return DeptCode;
        }



        static public void createMissingRelatedParties(Form oForm)
        {
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralService objEntryObject = null;
            SAPbobsCOM.GeneralData objEntryInfo = null;
            SAPbobsCOM.GeneralData objEntryLinesInfo = null;
            SAPbobsCOM.GeneralDataCollection objEntryLinesObject = null;
            SAPbouiCOM.DataTable DTResult = null;
            SAPbouiCOM.Matrix oMatriz = null;
            System.Data.DataTable oDT = B1.Base.UIOperations.FormsOperations.SapDataTableToDotNetDataTable(oForm.DataSources.DataTables.Item("DT_RP").SerializeAsXML(BoDataTableXmlSelect.dxs_All));
            int count = 1;
            DTResult = oForm.DataSources.DataTables.Item("DT_RES");

            oMatriz = (Matrix)oForm.Items.Item("Item_7").Specific;
            oMatriz.Columns.Item("LineId").DataBind.Bind("DT_RES", "LineId");
            oMatriz.Columns.Item("Col_0").DataBind.Bind("DT_RES", "CardCode");
            oMatriz.Columns.Item("Col_1").DataBind.Bind("DT_RES", "CardName");
            oMatriz.Columns.Item("Col_2").DataBind.Bind("DT_RES", "Success");
            oMatriz.Columns.Item("Col_3").DataBind.Bind("DT_RES", "Message");

            var strSQL = Queries.Instance.Queries().Get("GetBPBy_licTradNum");
            var objRecordSet = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                objCompanyService = MainObject.Instance.B1Company.GetCompanyService();

                foreach (DataRow dtr in oDT.Rows)
                {
                    try
                    {
                        objEntryObject = objCompanyService.GetGeneralService("HCO_FRP1100");
                    objEntryInfo = (GeneralData)objEntryObject.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);
                    objEntryInfo.SetProperty("Code", dtr["CardCode"]);
                    objEntryInfo.SetProperty("Name", dtr["CardName"]);
                    objEntryInfo.SetProperty("U_CardType", dtr["CardType"]);
                    objEntryInfo.SetProperty("U_CardTypeID", dtr["CardTypeC"]);
                    objEntryInfo.SetProperty("U_DocType", dtr["DocType"]);
                    objEntryInfo.SetProperty("U_DocTypeID", dtr["DocTypeC"]);
                    if (dtr["LicTradNum"].ToString().Contains("-"))
                    {
                        var lictrd = dtr["LicTradNum"].ToString().Split('-');
                        objEntryInfo.SetProperty("U_LicTradNum", lictrd[0]);
                        objEntryInfo.SetProperty("U_AuthDig", lictrd[1]);
                    }
                    else
                    {
                        objEntryInfo.SetProperty("U_LicTradNum", dtr["LicTradNum"].ToString());
                    }

                    objEntryInfo.SetProperty("U_FirstName", dtr["FirstName"]);
                    objEntryInfo.SetProperty("U_MiddleName", dtr["MiddleName"]);
                    objEntryInfo.SetProperty("U_LastName", dtr["LastName"]);
                    objEntryInfo.SetProperty("U_ScndSrName", dtr["ScndSrName"]);
                    objEntryInfo.SetProperty("U_AddNames", dtr["AddNames"]);
                    objEntryInfo.SetProperty("U_Country", dtr["Country"]);
                    objEntryInfo.SetProperty("U_TaxCardType", dtr["TaxCardType"]);
                    objEntryInfo.SetProperty("U_TaxCardTypeID", dtr["TaxCardTypeC"]);
                    objEntryInfo.SetProperty("U_DeptCode", dtr["DeptC"]);
                    objEntryInfo.SetProperty("U_TaxRegim", dtr["TaxRegim"]);
                    objEntryInfo.SetProperty("U_TaxRegimID", dtr["TaxRegimC"]);
                    objEntryInfo.SetProperty("U_CountyCode", dtr["CountyC"]);
                    objEntryInfo.SetProperty("U_EconActvty", dtr["EconActvty"]);
                    objEntryInfo.SetProperty("U_MainAddress", dtr["MainAddress"]);
                    objEntryInfo.SetProperty("U_Phone1", dtr["Phone1"]);
                    objEntryInfo.SetProperty("U_Phone2", dtr["Phone2"]);
                    objEntryInfo.SetProperty("U_Email", dtr["Email"]);
                    objEntryInfo.SetProperty("U_ZipCode", dtr["ZipCode"]);
                    objEntryInfo.SetProperty("U_CountryCode", dtr["CountryC"]);
                    objEntryInfo.SetProperty("U_CountyName", dtr["County"]);

                    objRecordSet.DoQuery(string.Format(strSQL, dtr["LicTradNum"].ToString()));
                    if (objRecordSet.RecordCount > 0)
                    {
                        objEntryLinesObject = objEntryInfo.Child("HCO_RP1101");
                        while (!objRecordSet.EoF)
                        {                            
                            objEntryLinesInfo = objEntryLinesObject.Add();
                            objEntryLinesInfo.SetProperty("U_CardType", objRecordSet.Fields.Item("CardType").Value.ToString());
                            objEntryLinesInfo.SetProperty("U_CardCode", objRecordSet.Fields.Item("CardCode").Value.ToString());
                            objEntryLinesInfo.SetProperty("U_CardName", objRecordSet.Fields.Item("CardName").Value.ToString());
                            objRecordSet.MoveNext();
                        }
                    }


                        objEntryObject.Add(objEntryInfo);
                        DTResult.Rows.Add(1);
                        DTResult.SetValue("LineId", count - 1, count);
                        DTResult.SetValue("CardCode", count - 1, dtr["CardCode"]);
                        DTResult.SetValue("CardName", count - 1, dtr["CardName"]);
                        DTResult.SetValue("Success", count - 1, "Si");
                        DTResult.SetValue("Message", count - 1, "Tercero registrado con éxito.");
                    }
                    catch(Exception ex)
                    {
                        DTResult.Rows.Add(1);
                        DTResult.SetValue("LineId", count - 1, count);
                        DTResult.SetValue("CardCode", count - 1, dtr["CardCode"]);
                        DTResult.SetValue("CardName", count - 1, dtr["CardName"]);
                        DTResult.SetValue("Success", count - 1, "No");
                        DTResult.SetValue("Message", count - 1, ex.Message);
                    }

                    count++;
                }
                oMatriz.LoadFromDataSourceEx();
                oMatriz.AutoResizeColumns();

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
                
                if (oForm != null)
                {
                    oForm.Freeze(false);
                }
            }
        }

        static public void ActualizarInfoCapitalizacion(string docEntry)
        {
            UpdateJournalCapitalization(docEntry);
            if (IsCheckItemCapitalizationMarked().Equals("Y"))
                UpdateItemsCapitalization(docEntry);
        }

        static public void ValorizationExecution(string docEntry)
        {
            var journal = (JournalEntries) MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oJournalEntries);
            var querySN = Queries.Instance.Queries().Get("GetDefaultSN");
            var record = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                record.DoQuery(querySN);
            var RelPartyCode = record.Fields.Item("U_DefaultSN").Value.ToString();
            var queryThird = string.Format(Queries.Instance.Queries().Get("GetValorizationExecution"), docEntry);
            record.DoQuery(queryThird);
            while(!record.EoF)
            {
                if( journal.GetByKey(int.Parse(record.Fields.Item("TransId").Value.ToString())) )
                {
                    for(int i=0; i<journal.Lines.Count; i++)
                    {
                        journal.Lines.SetCurrentLine(i);
                        journal.Lines.UserFields.Fields.Item("U_HCO_RELPAR").Value = RelPartyCode;
                    }

                    var resp = journal.Update();
                }

                record.MoveNext();
            }
        }

        static private string IsCheckItemCapitalizationMarked()
        {
            var strSQL = Queries.Instance.Queries().Get("GetCheckItemAF");
            var objRecordSet = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            objRecordSet.DoQuery(strSQL);
            objRecordSet.MoveFirst();

            if (objRecordSet.RecordCount > 0)
            {
                if (objRecordSet.Fields.Item("U_DesmCompAF").Value.Equals("Y"))
                    return "Y";
            }

            return "N";
        }

        static public void SetValueCapitalizacion(string docEntry, string type)
        {
            var strSQL = string.Format(Queries.Instance.Queries().Get("CheckCapitalization"), type, docEntry);
            var objRecordSet = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            objRecordSet.DoQuery(strSQL);
            objRecordSet.MoveFirst();
            if (objRecordSet.RecordCount > 0)
            {
                ActualizarInfoCapitalizacion(objRecordSet.Fields.Item("DocEntry").Value.ToString());
            }
        }

        static public void SetCapitalizationNC(string formUid, string docEntry, string type)
        {
            try
            {
                var form = MainObject.Instance.B1Application.Forms.Item(formUid);
                var strSQL = string.Format(Queries.Instance.Queries().Get("CheckContainsAsset"), "RPC1", docEntry);
                var objRecordSet = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                objRecordSet.DoQuery(strSQL);
                objRecordSet.MoveFirst();
                if (objRecordSet.RecordCount > 0)
                {
                    CreateCapitalizationNC(docEntry);
                }
            }
            catch(Exception ex)
            {

            }
        }

        static private void CreateCapitalizationNC(string docEntry)
        {
            var item = (SAPbobsCOM.Items)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oItems);
            var ncP = (SAPbobsCOM.Documents)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oPurchaseCreditNotes);
            if( ncP.GetByKey(int.Parse(docEntry)) )
            {
                var assetServices = (AssetDocumentService)MainObject.Instance.B1Company.GetCompanyService().GetBusinessService(ServiceTypes.AssetCapitalizationCreditMemoService);
                var faDocumentParams = (AssetDocument)assetServices.GetDataInterface(AssetDocumentServiceDataInterfaces.adsAssetDocument);
                    faDocumentParams.Reference = ncP.DocEntry.ToString();
                for (int i = 0; i < ncP.Lines.Count; i++)
                {
                    ncP.Lines.SetCurrentLine(i);
                    if (item.GetByKey(ncP.Lines.ItemCode))
                    {
                        if (item.ItemType == ItemTypeEnum.itFixedAssets)
                        {
                            var line = faDocumentParams.AssetDocumentLineCollection.Add();
                            line.AssetNumber = ncP.Lines.ItemCode;
                            line.TotalLC = ncP.Lines.LineTotal;
                        }
                    }
                }

                var response = assetServices.Add(faDocumentParams);
            }
        }

        static private void UpdateJournalCapitalization(string docEntry)
        {
            var assetServices = (AssetDocumentService)MainObject.Instance.B1Company.GetCompanyService().GetBusinessService(ServiceTypes.AssetCapitalizationService);
            var faDocumentParams = (AssetDocumentParams)assetServices.GetDataInterface(AssetDocumentServiceDataInterfaces.adsAssetDocumentParams);
                faDocumentParams.Code = int.Parse(docEntry);
            var AssetDocument = assetServices.Get(faDocumentParams);
            var journal = (JournalEntries)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oJournalEntries);

            for (int i = 0; i < AssetDocument.AssetDocumentAreaJournalCollection.Count; i++)
            {
                if (journal.GetByKey(AssetDocument.AssetDocumentAreaJournalCollection.Item(i).TransactionNumber))
                {
                    journal.UserFields.Fields.Item("U_HCO_ValAre").Value = GetValueDepreciationArea(AssetDocument.AssetDocumentAreaJournalCollection.Item(i).DepreciationArea);
                    journal.Update();
                }
            }
        }

        static private void UpdateItemsCapitalization(string docEntry)
        {
            var assetServices = (AssetDocumentService)MainObject.Instance.B1Company.GetCompanyService().GetBusinessService(ServiceTypes.AssetCapitalizationService);
            var faDocumentParams = (AssetDocumentParams)assetServices.GetDataInterface(AssetDocumentServiceDataInterfaces.adsAssetDocumentParams);
            faDocumentParams.Code = int.Parse(docEntry);

            var AssetDocument = assetServices.Get(faDocumentParams);
            var item = (SAPbobsCOM.Items)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oItems);

            for (int i = 0; i < AssetDocument.AssetDocumentLineCollection.Count; i++)
            {
                var itemCode = AssetDocument.AssetDocumentLineCollection.Item(i).AssetNumber;
                if (item.GetByKey(itemCode))
                {
                    item.PurchaseItem = BoYesNoEnum.tNO;
                    item.Update();
                }
            }
        }

        static private string GetValueDepreciationArea(string value)
        {
            var strSQL = string.Format(Queries.Instance.Queries().Get("GetValueDep"), value);
            var objRecordSet = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            objRecordSet.DoQuery(strSQL);
            objRecordSet.MoveFirst();

            if (objRecordSet.RecordCount > 0)
                return objRecordSet.Fields.Item("ValueDep").Value.ToString();

            return string.Empty;
        }

        static public void LoadDataThird(SAPbouiCOM.BusinessObjectInfo pVal)
        {
            var form = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            var cardCode = form.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0);
            var valueThird = GetValueThird(cardCode);
            form.DataSources.UserDataSources.Item("UD_RelPty").Value = valueThird;
        }
        static public string GetValueThird(string cardCode)
        {
            var thrid = string.Empty;
            var strSQL = string.Format(Queries.Instance.Queries().Get("GetThirdRelated"), cardCode);
            var objRecordSet = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            objRecordSet.DoQuery(strSQL);
            objRecordSet.MoveFirst();

            if (objRecordSet.RecordCount > 0)
            {
                thrid = objRecordSet.Fields.Item("Code").Value.ToString();
            }

            return thrid;
        }

        static public void UpdateJournalPaymentCreated(string objType, string document)
        {
            var oDoc = (SAPbobsCOM.Payments)MainObject.Instance.B1Company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), objType));

            try
            {
                var RelPartyCode = string.Empty;
                var xml = new XmlDocument();
                xml.LoadXml(document);
                oDoc.GetByKey(Int32.Parse(xml.InnerText));

                var query = string.Format(Queries.Instance.Queries().Get("GetTransId"), objType, xml.InnerText);
                var record = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                record.DoQuery(query);

                var transId = (string)record.Fields.Item("TransId").Value.ToString();
                var cardCodeDoc = oDoc.CardCode;

                if (Array.IndexOf(Parameters.docTercDef, objType) >= 0)
                {
                    var querySN = Queries.Instance.Queries().Get("GetDefaultSN");
                    record.DoQuery(querySN);

                    RelPartyCode = record.Fields.Item("U_DefaultSN").Value.ToString();
                }
                else
                {
                    RelPartyCode = GetValueThird(cardCodeDoc);
                }

                if (transId.Equals(string.Empty) || RelPartyCode.Equals(string.Empty))
                    return;

                var journal = (JournalEntries)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oJournalEntries);
                if (journal.GetByKey(int.Parse(transId)))
                {
                    for (int i = 0; i < journal.Lines.Count; i++)
                    {
                        journal.Lines.SetCurrentLine(i);
                        journal.Lines.UserFields.Fields.Item("U_HCO_RELPAR").Value = RelPartyCode;
                    }

                    var resp = journal.Update();
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
        static public void UpdateJournalDocumentCreated(string objType, string document)
        {
            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)MainObject.Instance.B1Company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), objType));

            try
            {
                var RelPartyCode = string.Empty;
                var xml = new XmlDocument();
                xml.LoadXml(document);
                oDoc.GetByKey(Int32.Parse(xml.InnerText));

                var query = string.Format(Queries.Instance.Queries().Get("GetTransId"), objType, xml.InnerText);
                var record = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                record.DoQuery(query);

                var transId = (string)record.Fields.Item("TransId").Value.ToString();
                var cardCodeDoc = oDoc.CardCode;

                if (Array.IndexOf(Parameters.docTercDef, objType) >= 0)
                {
                    var querySN = Queries.Instance.Queries().Get("GetDefaultSN");
                    record.DoQuery(querySN);

                    RelPartyCode = record.Fields.Item("U_DefaultSN").Value.ToString();
                }
                else
                {
                    RelPartyCode = GetValueThird(cardCodeDoc);
                    var queryPayment = string.Format(Queries.Instance.Queries().Get("GetReceiptFromInvoice"), objType, xml.InnerText);
                    var recordPayment = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                        recordPayment.DoQuery(queryPayment);

                    if( recordPayment.RecordCount > 0 )
                    {
                        if (!recordPayment.Fields.Item("Receipt").Value.ToString().Equals(string.Empty))                                                                    
                        {
                            try 
                            {
                                var paym = (Payments)MainObject.Instance.B1Company.GetBusinessObject((BoObjectTypes)Enum.Parse(typeof(BoObjectTypes), recordPayment.Fields.Item("ObjType").Value.ToString()));
                                if( paym.GetByKey(int.Parse(recordPayment.Fields.Item("Receipt").Value.ToString())) )
                                {
                                    var journalPay = (JournalEntries)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oJournalEntries);
                                    if (journalPay.GetByKey(int.Parse(recordPayment.Fields.Item("PaymentJournal").Value.ToString()))) 
                                    {
                                        for (int i = 0; i < journalPay.Lines.Count; i++)
                                        {
                                            journalPay.Lines.SetCurrentLine(i);
                                            journalPay.Lines.UserFields.Fields.Item("U_HCO_RELPAR").Value = RelPartyCode;
                                        }

                                        var resp = journalPay.Update();
                                    }
                                }
                            }
                            catch
                            {

                            }
                        }
                    }
                }

                if (transId.Equals(string.Empty) || RelPartyCode.Equals(string.Empty))
                    return;

                var journal = (JournalEntries)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oJournalEntries);
                if (journal.GetByKey(int.Parse(transId)))
                {
                    for (int i = 0; i < journal.Lines.Count; i++)
                    {
                        journal.Lines.SetCurrentLine(i);
                        journal.Lines.UserFields.Fields.Item("U_HCO_RELPAR").Value = RelPartyCode;
                    }

                    var resp = journal.Update();
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
        static public void UpdateJournalPayment(BusinessObjectInfo pVal)
        {
            try
            {
                var journal = (SAPbobsCOM.JournalEntries) MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oJournalEntries);
                var xmlDoc = new XmlDocument();
                    xmlDoc.LoadXml(pVal.ObjectKey);

                var form = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                var strSQL = string.Format(Queries.Instance.Queries().Get("GetJournalPaymentNumber"), form.DataSources.DBDataSources.Item(0).TableName, xmlDoc.InnerText);
                var objRS = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    objRS.DoQuery(strSQL);

                if (journal.GetByKey(int.Parse(objRS.Fields.Item("TransId").Value.ToString())))
                {
                    var queryUpdate = string.Format(Queries.Instance.Queries().Get("CheckPaymentAccount"), form.DataSources.DBDataSources.Item(0).TableName, form.DataSources.DBDataSources.Item(0).TableName == "OVPM" ? "VPM4" : "RCT4", objRS.Fields.Item("TransId").Value);
                    objRS.DoQuery(queryUpdate);

                    while (!objRS.EoF)
                    {
                        journal.Lines.SetCurrentLine(int.Parse(objRS.Fields.Item("Line_ID").Value.ToString()));
                        journal.Lines.UserFields.Fields.Item("U_HCO_RELPAR").Value = objRS.Fields.Item("Tercero").Value;

                        objRS.MoveNext();
                    }

                    var resp = journal.Update();
                    var msg = MainObject.Instance.B1Company.GetLastErrorDescription();
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
        static public void loadMissingRelatedPartiesForm()
        {
            string strSQL = "";
            SAPbobsCOM.Recordset objRS = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.FormCreationParams objParams = null;
            SAPbouiCOM.DataTable objDT = null;
            //SAPbouiCOM.Item objItem = null;
            SAPbouiCOM.Grid objGrid = null;
            SAPbouiCOM.GridColumn objGridColumn = null;
            SAPbouiCOM.EditTextColumn oEditTExt = null;


            try
            {
                objParams = (FormCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                objParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                objParams.XmlData = "";//RelatedPartiesRes.HCO_Terceros_Relacionados_Faltantes;
                objParams.FormType = "HCO_FTRA1";
                objForm = MainObject.Instance.B1Application.Forms.AddEx(objParams);
                objDT = objForm.DataSources.DataTables.Item("DT_TRA");

                strSQL = Queries.Instance.Queries().Get("GetMissingRP");
                objRS = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objDT.ExecuteQuery(strSQL);

                #region Format Grid
                objGrid = (Grid)objForm.Items.Item("grTRA").Specific;

                //objGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;


                objGridColumn = objGrid.Columns.Item(0);
                objGridColumn.Editable = false;

                objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oEditTExt = (SAPbouiCOM.EditTextColumn)objGridColumn;
                oEditTExt.LinkedObjectType = "2";


                objGridColumn = objGrid.Columns.Item(1);
                objGridColumn.Editable = false;


                objGridColumn = objGrid.Columns.Item(2);
                objGridColumn.Editable = false;

                objGrid.AutoResizeColumns();


                #endregion


                objForm.Visible = true;

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
        public static void SetChooseFromList(ItemEvent pVal)
        {
            var oForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            var field = ((EditText)oForm.Items.Item(pVal.ItemUID).Specific).DataBind.Alias;

            switch (pVal.ItemUID)
            {
                case "Item_11":
                    if (B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Name")[0].ToString().Equals(string.Empty)) return;
                    switch (pVal.FormTypeEx)
                    {
                        case "HCO_FRP0001":
                            oForm.DataSources.DBDataSources.Item(0).SetValue(field, 0, B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Code")[0].ToString());
                            oForm.DataSources.DBDataSources.Item(0).SetValue("U_CardName", 0, B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Name")[0].ToString());
                            break;
                        case "HCO_FRP1100":
                            oForm.DataSources.DBDataSources.Item(0).SetValue(field, 0, B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Name")[0].ToString());
                            oForm.DataSources.DBDataSources.Item(0).SetValue("U_CardTypeID", 0, B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Code")[0].ToString());
                            break;
                    }                                                           
                    break;
                case "Item_26":
                    if (B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Name")[0].ToString().Equals(string.Empty)) return;
                    oForm.DataSources.DBDataSources.Item(0).SetValue(field, 0, B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Name")[0].ToString());
                    oForm.DataSources.DBDataSources.Item(0).SetValue("U_DocTypeID", 0, B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Code")[0].ToString());
                    break;
                case "Item_35":
                    if (B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Name")[0].ToString().Equals(string.Empty)) return;
                    oForm.DataSources.DBDataSources.Item(0).SetValue(field, 0, B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Name")[0].ToString());
                    oForm.DataSources.DBDataSources.Item(0).SetValue("U_TaxCardTypeID", 0, B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Code")[0].ToString());
                    break;
                case "Item_37":
                    if (B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Name")[0].ToString().Equals(string.Empty)) return;
                    oForm.DataSources.DBDataSources.Item(0).SetValue(field, 0, B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Name")[0].ToString());
                    oForm.DataSources.DBDataSources.Item(0).SetValue("U_TaxRegimID", 0, B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Code")[0].ToString());
                    break;
                case "Item_41":
                    if (B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Code")[0].ToString().Equals(string.Empty)) return;
                    oForm.DataSources.DBDataSources.Item(0).SetValue(field, 0, B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Code")[0].ToString());
                    break;
                case "Item_46":
                    if (B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Name")[0].ToString().Equals(string.Empty)) return;
                    oForm.DataSources.DBDataSources.Item(0).SetValue(field, 0, B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Name")[0].ToString());
                    oForm.DataSources.DBDataSources.Item(0).SetValue("U_CountyCode", 0, B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Code")[0].ToString());
                    oForm.DataSources.DBDataSources.Item(0).SetValue("U_DeptCode", 0, B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "U_Departamento")[0].ToString());
                    break;
                case "Item_48":
                    if (B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Name")[0].ToString().Equals(string.Empty)) return;
                    oForm.DataSources.DBDataSources.Item(0).SetValue(field, 0, B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Name")[0].ToString());
                    oForm.DataSources.DBDataSources.Item(0).SetValue("U_CountryCode", 0, B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Code")[0].ToString());
                    break;
                case "Item_42":
                    if (B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "CardCode")[0].ToString().Equals(string.Empty)) return;
                    oForm.DataSources.UserDataSources.Item("UD_SocD").Value = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "CardCode")[0].ToString();
                    break;
                case "Item_44":
                    if (B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "CardCode")[0].ToString().Equals(string.Empty)) return;
                    oForm.DataSources.UserDataSources.Item("UD_SocH").Value = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "CardCode")[0].ToString();
                    break;
                case "Item_47":
                    if (B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "AcctCode")[0].ToString().Equals(string.Empty)) return;
                    oForm.DataSources.UserDataSources.Item("UD_CtaD").Value = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "AcctCode")[0].ToString();
                    break;
                case "Item_49":
                    if (B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "AcctCode")[0].ToString().Equals(string.Empty)) return;
                    oForm.DataSources.UserDataSources.Item("UD_CtaH").Value = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "AcctCode")[0].ToString();
                    break;
                case "Item_63":
                    if (B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "OcrCode")[0].ToString().Equals(string.Empty)) return;
                    oForm.DataSources.UserDataSources.Item("UD_Dim1").Value = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "OcrCode")[0].ToString();
                    break;
                case "Item_64":
                    if (B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "OcrCode")[0].ToString().Equals(string.Empty)) return;
                    oForm.DataSources.UserDataSources.Item("UD_Dim2").Value = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "OcrCode")[0].ToString();
                    break;
                case "Item_65":
                    if (B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "OcrCode")[0].ToString().Equals(string.Empty)) return;
                    oForm.DataSources.UserDataSources.Item("UD_Dim3").Value = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "OcrCode")[0].ToString();
                    break;
                case "Item_66":
                    if (B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "OcrCode")[0].ToString().Equals(string.Empty)) return;
                    oForm.DataSources.UserDataSources.Item("UD_Dim4").Value = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "OcrCode")[0].ToString();
                    break;
                case "Item_0":
                    if (B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "OcrCode")[0].ToString().Equals(string.Empty)) return;
                    oForm.DataSources.UserDataSources.Item("UD_Dim5").Value = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "OcrCode")[0].ToString();
                    break;
                case "txtThird":
                    if (B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Code")[0].ToString().Equals(string.Empty)) return;
                    ((EditText)oForm.Items.Item(pVal.ItemUID).Specific).Value = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Code")[0].ToString();
                    break;
            }

            if (oForm.Mode == BoFormMode.fm_OK_MODE)
                oForm.Mode = BoFormMode.fm_UPDATE_MODE;
        }
        public static bool ValidateFields(ItemEvent pVal)
        {
            var oForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            if (oForm.Mode != BoFormMode.fm_FIND_MODE)
            {
                var ThirdParty_type = oForm.DataSources.DBDataSources.Item(0).GetValue("U_TaxCardTypeID", 0).Trim();

                if (oForm.DataSources.DBDataSources.Item(0).GetValue("Code", 0).Trim().Equals(string.Empty))
                {
                    MainObject.Instance.B1Application.SetStatusBarMessage("Falta código del tercero relacionado.");
                    oForm.Items.Item("0_U_E").Click(BoCellClickType.ct_Regular);
                    return false;
                }

                if (oForm.DataSources.DBDataSources.Item(0).GetValue("Name", 0).Trim().Equals(string.Empty))
                {
                    MainObject.Instance.B1Application.SetStatusBarMessage("Falta razón social del tercero.");
                    oForm.Items.Item("Item_7").Click(BoCellClickType.ct_Regular);
                    return false;
                }

                if (oForm.DataSources.DBDataSources.Item(0).GetValue("U_CardType", 0).Trim().Equals(string.Empty))
                {
                    MainObject.Instance.B1Application.SetStatusBarMessage("Falta tipo de tercero.");
                    oForm.Items.Item("Item_11").Click(BoCellClickType.ct_Regular);
                    return false;
                }

                if (oForm.DataSources.DBDataSources.Item(0).GetValue("U_DocType", 0).Trim().Equals(string.Empty))
                {
                    MainObject.Instance.B1Application.SetStatusBarMessage("Falta tipo de documento.");
                    oForm.Items.Item("Item_26").Click(BoCellClickType.ct_Regular);
                    return false;
                }

                if (oForm.DataSources.DBDataSources.Item(0).GetValue("U_LicTradNum", 0).Trim().Equals(string.Empty))
                {
                    MainObject.Instance.B1Application.SetStatusBarMessage("Falta número de identificación.");
                    oForm.Items.Item("Item_13").Click(BoCellClickType.ct_Regular);
                    return false;
                }

                if (ThirdParty_type.Equals("1"))
                {
                    if (oForm.DataSources.DBDataSources.Item(0).GetValue("U_FirstName", 0).Trim().Equals(string.Empty))
                    {
                        MainObject.Instance.B1Application.SetStatusBarMessage("Falta primer nombre.");
                        oForm.Items.Item("Item_31").Click(BoCellClickType.ct_Regular);
                        return false;
                    }

                    if (oForm.DataSources.DBDataSources.Item(0).GetValue("U_LastName", 0).Trim().Equals(string.Empty))
                    {
                        MainObject.Instance.B1Application.SetStatusBarMessage("Falta primer apellido.");
                        oForm.Items.Item("Item_32").Click(BoCellClickType.ct_Regular);
                        return false;
                    }
                }
                if (oForm.DataSources.DBDataSources.Item(0).GetValue("U_Country", 0).Trim().Equals(string.Empty))
                {
                    MainObject.Instance.B1Application.SetStatusBarMessage("Falta país del tercero.");
                    oForm.Items.Item("Item_48").Click(BoCellClickType.ct_Regular);
                    return false;
                }

                if (oForm.DataSources.DBDataSources.Item(0).GetValue("U_TaxCardTypeID", 0).Trim().Equals(string.Empty))
                {
                    MainObject.Instance.B1Application.SetStatusBarMessage("Falta tipo de contribuyente.");
                    oForm.Items.Item("Item_35").Click(BoCellClickType.ct_Regular);
                    return false;
                }

                if (oForm.DataSources.DBDataSources.Item(0).GetValue("U_CountyCode", 0).Trim().Equals(string.Empty))
                {
                    MainObject.Instance.B1Application.SetStatusBarMessage("Falta municipio del tercero.");
                    oForm.Items.Item("Item_46").Click(BoCellClickType.ct_Regular);
                    return false;
                }

                if (oForm.DataSources.DBDataSources.Item(0).GetValue("U_MainAddress", 0).Trim().Equals(string.Empty))
                {
                    MainObject.Instance.B1Application.SetStatusBarMessage("Falta dirección del tercero.");
                    oForm.Items.Item("Item_47").Click(BoCellClickType.ct_Regular);
                    return false;
                }
            }
            return true;
        }

        public static bool ValidateFieldInventoryTransfer(ItemEvent pVal)
        {
            var oForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            if (oForm.Mode != BoFormMode.fm_FIND_MODE)
            {
                var cardCode = oForm.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0).Trim();
                if( !string.IsNullOrEmpty(cardCode) )
                {
                    var bp = (BusinessPartners) MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                    if( bp.GetByKey(cardCode) )
                    {
                        var thirdPartyQuery = string.Format(Queries.Instance.Queries().Get("CheckRelThird"), cardCode);
                        var record = (SAPbobsCOM.Recordset) MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                        record.DoQuery(thirdPartyQuery);
                        if (record.RecordCount == 0) 
                        {
                            MainObject.Instance.B1Application.SetStatusBarMessage("Falta código del tercero relacionado.");
                            return false;
                        }
                    }
                }
            }
            return true;
        }

        public static bool ValidateFieldsPayment(ItemEvent pVal)
        {
            var oForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            if (oForm.Mode != BoFormMode.fm_FIND_MODE)
            {
                var docType = oForm.DataSources.DBDataSources.Item(0).GetValue("DocType", 0).Trim();

                if (docType.Equals("A"))
                {
                    var matrix = (SAPbouiCOM.Matrix)oForm.Items.Item("71").Specific;
                    var thirdParty = oForm.DataSources.DBDataSources.Item(0).GetValue("U_HCO_RELPAR", 0).Trim();
                    if (thirdParty.Equals(string.Empty))
                    {
                        MainObject.Instance.B1Application.SetStatusBarMessage("Falta código del tercero relacionado.");
                        return false;
                    }

                    for(int i=1; i<=matrix.RowCount; i++)
                    {
                        if( ((EditText)matrix.GetCellSpecific("U_HCO_RELPAR", i)).Value.Equals(string.Empty) && !((EditText)matrix.GetCellSpecific("8", i)).Value.Equals(string.Empty))
                        {
                            MainObject.Instance.B1Application.SetStatusBarMessage($"Falta código del tercero relacionado en la linea {i}.");
                            return false;
                        }
                    }
                }
            }
            return true;
        }

        public static void SetChooseFromListMatrix(ItemEvent pVal)
        {
            var oForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            var matrix = ((Matrix)oForm.Items.Item(pVal.ItemUID).Specific);
            var cardCode = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "CardCode")[0].ToString();
            var cardType = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "CardType")[0].ToString();
            var cardName = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "CardName")[0].ToString();

            if(validateCardCodeInMatrix(cardCode, oForm))
            {
                MainObject.Instance.B1Application.MessageBox("El " + (cardType.Equals("C") ? "cliente " : "proveedor ") + "seleccionado ya se encuentra registrado previamente.");
                return;
            }

            if (cardCode.Equals(string.Empty))
                return;

            if (!ValidateLicTradNum(cardCode, oForm))
            {
                MainObject.Instance.B1Application.MessageBox("El " + (cardType.Equals("C") ? "cliente " : "proveedor ") + "seleccionado debe corresponder al mismo número de identificación del tercero relacionado.");
                return;
            }

            switch (pVal.ItemUID)
            {
                case "Item_52":
                    matrix.SetCellWithoutValidation(pVal.Row, "Col_0", cardCode);
                    matrix.SetCellWithoutValidation(pVal.Row, "Col_1", cardName);
                    matrix.SetCellWithoutValidation(pVal.Row, "Col_3", cardType);
                    break;
            }

            if (oForm.Mode == BoFormMode.fm_OK_MODE)
                oForm.Mode = BoFormMode.fm_UPDATE_MODE;
        }

        public static void SetChooseFromListThirdPayment(ItemEvent pVal)
        {
            var oForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            var third = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Code")[0].ToString();
            if (third.Equals(string.Empty))
                return;

            if (!pVal.ColUID.Equals(string.Empty))
            {
                var matrix = ((Matrix)oForm.Items.Item(pVal.ItemUID).Specific);
                matrix.SetCellWithoutValidation(pVal.Row, pVal.ColUID, third);
            }
            else
            {
                try
                {
                    ((EditText)oForm.Items.Item(pVal.ItemUID).Specific).Value = third;
                }
                finally { }
            }
        }

        public static void AddFieldsJournalChangesTax(ItemEvent pVal)
        {
            var oForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                oForm.DataSources.UserDataSources.Add("UD_MetVat", BoDataType.dt_SHORT_TEXT, 200);

            var labelReference = oForm.Items.Item("29");
            var itemReference = oForm.Items.Item("28");
            var labelAdd = oForm.Items.Add("lblMet", BoFormItemTypes.it_STATIC);
            var comboAdd = oForm.Items.Add("itmMet", BoFormItemTypes.it_COMBO_BOX);

            comboAdd.Top = itemReference.Top;
            comboAdd.Left = itemReference.Left;
            comboAdd.DisplayDesc = true;
            ((ComboBox)comboAdd.Specific).DataBind.SetBound(true, "", "UD_MetVat");
            ((ComboBox)comboAdd.Specific).ValidValues.Add("C", "Común");
            ((ComboBox)comboAdd.Specific).ValidValues.Add("I", "IFRS");
            ((ComboBox)comboAdd.Specific).ValidValues.Add("L", "Local");

            labelAdd.Width = 100;
            labelAdd.Top = labelReference.Top;
            labelAdd.Left = labelReference.Left;
            ((StaticText)labelAdd.Specific).Caption = "Área de valorización";
        }

        public static void UpdateJournalChangesTax(BusinessObjectInfo BusinessObjectInfo)
        {
            var oForm = MainObject.Instance.B1Application.Forms.Item(BusinessObjectInfo.FormUID);
            var journal = (SAPbobsCOM.JournalEntries)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oJournalEntries);
            var oRS = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            var hash = ((EditText)oForm.Items.Item("4").Specific).Value;
            var areaVal = oForm.DataSources.UserDataSources.Item("UD_MetVat").Value;
            var strSQL = string.Format(Queries.Instance.Queries().Get("GetChangesDifferences"), hash, (BusinessObjectInfo.FormTypeEx.Equals("369") ? "HDCA" : "DCO"));
                oRS.DoQuery(strSQL);

            while(!oRS.EoF)
            {
                var third = oRS.Fields.Item("RelPar").Value.ToString();
                var transId = int.Parse(oRS.Fields.Item("TransId").Value.ToString());

                if(journal.GetByKey(transId))
                {
                    for(int i=0; i<journal.Lines.Count; i++)
                    {
                        journal.Lines.SetCurrentLine(i);
                        journal.Lines.UserFields.Fields.Item("U_HCO_RELPAR").Value = third;
                    }
                }

                journal.UserFields.Fields.Item("U_HCO_ValAre").Value = areaVal;
                var resp = journal.Update();
                oRS.MoveNext();
            }
        }

        private static bool ValidateLicTradNum(string cardCode, Form oForm)
        {
            try
            {
                string LicTradNumRP = oForm.DataSources.DBDataSources.Item("@HCO_RP1100").GetValue("U_LicTradNum", 0) + "-" + oForm.DataSources.DBDataSources.Item("@HCO_RP1100").GetValue("U_AuthDig", 0);
                BusinessPartners oBP = (BusinessPartners)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                oBP.GetByKey(cardCode);

                return LicTradNumRP.Replace("-","").ToString().Equals(oBP.FederalTaxID.Replace("-",""));
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);
                return false;

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                return false;
            }
        }

        private static bool validateCardCodeInMatrix(string cardCode, Form oForm)
        {
            try
            {
                var dataTable = B1.Base.UIOperations.FormsOperations.SapDBDataSourceToDotNetDataTable(oForm.DataSources.DBDataSources.Item("@HCO_RP1101"));
                return dataTable.AsEnumerable().Any(r => r.Field<string>("U_CardCode") == cardCode);
                
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);
                return false;

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                return false;
            }
        }

        public static void OpenThirdForm(ItemEvent pVal)
        {
            var oForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            var valueThird = oForm.DataSources.DBDataSources.Item(0).GetValue("U_HCO_RELPAR", 0);
            MainObject.Instance.B1Application.Menus.Item("HCO_MRP0009").Activate();
            var activeForm = MainObject.Instance.B1Application.Forms.GetForm("HCO_FRP1100", GetCountForm("HCO_FRP1100"));
            
            try
            {
                activeForm.Freeze(true);
                MainObject.Instance.B1Application.Menus.Item("1281").Activate();
                ((EditText)activeForm.Items.Item("0_U_E").Specific).Value = valueThird;
                activeForm.Items.Item("1").Click();
            }
            finally
            {
                activeForm.Freeze(false);
            }
        }

        private static int GetCountForm(string type)
        {
            var cantidad = 0;
            var cantForm = MainObject.Instance.B1Application.Forms.GetEnumerator();
            for (int i = 0; i < MainObject.Instance.B1Application.Forms.Count; i++) 
            {
                if (MainObject.Instance.B1Application.Forms.Item(i).TypeEx == type)
                    cantidad++;
            }
            return cantidad;
        }
    }
}
