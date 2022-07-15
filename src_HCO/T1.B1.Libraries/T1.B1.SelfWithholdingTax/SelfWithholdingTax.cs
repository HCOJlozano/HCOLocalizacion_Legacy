using System;
using System.Collections.Generic;
using log4net;
using System.Runtime.InteropServices;
using System.Xml;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Linq;

namespace T1.B1.SelfWithholdingTax
{
    public class SelfWithholdingTax
    {
        private static SelfWithholdingTax objWithHoldingTax;
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);

        private SelfWithholdingTax()
        {
            if (objWithHoldingTax == null)
            {
                objWithHoldingTax = new SelfWithholdingTax();
            }
        }

        #region Add SWT normal operation

        public List<SelfWithholdingTaxTransaction> getWTTransaction(int DocEntry)
        {
            List<SelfWithholdingTaxTransaction> objResult = new List<SelfWithholdingTaxTransaction>();
            SAPbobsCOM.GeneralService objService = null;
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralDataParams objParams = null;
            SAPbobsCOM.GeneralData objGeneralData = null;

            try
            {
                objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                objService = objCompanyService.GetGeneralService(Settings._SelfWithHoldingTax.SWtaxUDOTransaction);
                objParams = (GeneralDataParams)objService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                objParams.SetProperty("DocEntry", DocEntry);
                objGeneralData = objService.GetByParams(objParams);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            return objResult;
        }

        public static string GetXmlUDO(Settings.SelfWithHoldingTaxTypes type)
        {
            XmlDocument oXML = new XmlDocument();
            switch (type)
            {
                case Settings.SelfWithHoldingTaxTypes.SELFWHITHHOLDINGTAX_CONFIGURAION:
                    oXML.Load(AppDomain.CurrentDomain.BaseDirectory + "\\Forms\\Autorretenciones.srf");
                    break;
                default:
                    return string.Empty;
            }
            return oXML.InnerXml;
        }

        public static void addSelfWithHoldingTax(SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            try
            {
                SAPbobsCOM.Documents objDoc = null;
                if (BusinessObjectInfo.Type == "13")
                {
                    objDoc = (Documents)MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                }
                else if (BusinessObjectInfo.Type == "14")
                {
                    objDoc = (Documents)MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
                }

                if (objDoc == null) return;
                if (objDoc.Browser.GetByKeys(BusinessObjectInfo.ObjectKey))
                {
                    calcSelfWTax(objDoc, BusinessObjectInfo.Type);
                }
                else
                {
                    _Logger.Error("Could not retrive Document with key " + BusinessObjectInfo.ObjectKey);
                    MainObject.Instance.B1Application.SetStatusBarMessage("T1: Could not retrive Document. Self WithHolding was not calculated");
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        private static SelfWothholdingTaxResult calcSelfWTax(SAPbobsCOM.Documents objDoc, string DocType)
        {
            SelfWothholdingTaxResult objResult = new SelfWothholdingTaxResult();
            try
            {
                List<SelfWithholdingTaxInfo> lSelfWithHolding = getSelfWithholdingTax(objDoc, DocType);
                bool blCancelation = false;
                string strThirdParty = "";

                if (DocType == "14")
                {
                    blCancelation = true;
                }

                if (lSelfWithHolding.Count > 0)
                {
                    strThirdParty = B1.RelatedParties.Instance.GetValueThird(objDoc.CardCode);

                    SAPbobsCOM.JournalEntries objJournal = (JournalEntries)MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                    objJournal.Memo = string.Concat(new object[] { "Autoretención para el documento ", objDoc.DocNum });
                    objJournal.Reference2 = objDoc.DocEntry.ToString();
                    objJournal.Reference = objDoc.DocNum.ToString();
                    objJournal.ReferenceDate = objDoc.DocDate;
                    objJournal.DueDate = objDoc.DocDueDate;
                    objJournal.TaxDate = objDoc.TaxDate;

                    if (Settings._SelfWithHoldingTax.WTaxTransCode.Length > 0)
                        objJournal.TransactionCode = Settings.WTaxTransCode;

                    foreach (SelfWithholdingTaxInfo sInfo in lSelfWithHolding)
                    {
                        if (DocType.Equals("13"))
                        {
                            objJournal.Lines.Credit = sInfo.dbWtAmount;
                            objJournal.Lines.AccountCode = !blCancelation ? sInfo.Credit : sInfo.Debit;
                            objJournal.Lines.Reference1 = objDoc.DocNum.ToString();
                            objJournal.Lines.Reference2 = objDoc.DocEntry.ToString();
                            objJournal.Lines.LineMemo = string.Concat(new object[] { "Autoretención de ", sInfo.Code, " para el documento ", objDoc.DocNum });
                            objJournal.Lines.UserFields.Fields.Item(Settings._SelfWithHoldingTax.relatedpartyFieldInLines).Value = strThirdParty;
                            objJournal.Lines.Add();

                            objJournal.Lines.Debit = sInfo.dbWtAmount;
                            objJournal.Lines.AccountCode = !blCancelation ? sInfo.Debit : sInfo.Credit;
                            objJournal.Lines.Reference1 = objDoc.DocNum.ToString();
                            objJournal.Lines.Reference2 = objDoc.DocEntry.ToString();
                            objJournal.Lines.LineMemo = string.Concat(new object[] { "Autoretención de ", sInfo.Code, " para el documento ", objDoc.DocNum });
                            objJournal.Lines.UserFields.Fields.Item(Settings._SelfWithHoldingTax.relatedpartyFieldInLines).Value = strThirdParty;
                            objJournal.Lines.Add();
                        }
                        else
                        {
                            objJournal.Lines.Debit = sInfo.dbWtAmount;
                            objJournal.Lines.AccountCode = !blCancelation ? sInfo.Credit : sInfo.Debit;
                            objJournal.Lines.Reference1 = objDoc.DocNum.ToString();
                            objJournal.Lines.Reference2 = objDoc.DocEntry.ToString();
                            objJournal.Lines.LineMemo = string.Concat(new object[] { "Autoretención de ", sInfo.Code, " para el documento ", objDoc.DocNum });
                            objJournal.Lines.UserFields.Fields.Item(Settings._SelfWithHoldingTax.relatedpartyFieldInLines).Value = strThirdParty;
                            objJournal.Lines.Add();

                            objJournal.Lines.Credit = sInfo.dbWtAmount;
                            objJournal.Lines.AccountCode = !blCancelation ? sInfo.Debit : sInfo.Credit;
                            objJournal.Lines.Reference1 = objDoc.DocNum.ToString();
                            objJournal.Lines.Reference2 = objDoc.DocEntry.ToString();
                            objJournal.Lines.LineMemo = string.Concat(new object[] { "Autoretención de ", sInfo.Code, " para el documento ", objDoc.DocNum });
                            objJournal.Lines.UserFields.Fields.Item(Settings._SelfWithHoldingTax.relatedpartyFieldInLines).Value = strThirdParty;
                            objJournal.Lines.Add();
                        }

                        lSelfWithHolding[lSelfWithHolding.IndexOf(sInfo)].DocEntry = objDoc.DocEntry;
                        lSelfWithHolding[lSelfWithHolding.IndexOf(sInfo)].DocNum = objDoc.DocNum;
                        lSelfWithHolding[lSelfWithHolding.IndexOf(sInfo)].CardCode = objDoc.CardCode;
                        lSelfWithHolding[lSelfWithHolding.IndexOf(sInfo)].DocType = objDoc.DocType.ToString();
                    }

                    if (objJournal.Add() == 0)
                    {
                        var journalCreated = MainObject.Instance.B1Company.GetNewObjectKey();
                        AddRegisterAutoWitholdingTax(objDoc, journalCreated, int.Parse(DocType));

                        MainObject.Instance.B1Application.SetStatusBarMessage("T1: Las autoretenciones se causaron con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        objResult.Message = "T1: Las autoretenciones se causaron con éxito.";
                        objResult.MessageCode = "";
                    }
                    else
                    {
                        string strMessage = MainObject.Instance.B1Company.GetLastErrorDescription();
                        _Logger.Error("Could not create SelfWithHolding Tax. " + strMessage);
                        MainObject.Instance.B1Application.SetStatusBarMessage("T1: Could not create SelfWithHolding Tax. Self WithHolding was not calculated." + strMessage);
                        objResult.Message = "T1: Could not create SelfWithHolding Tax. Self WithHolding was not calculated." + strMessage;
                        objResult.MessageCode = MainObject.Instance.B1Company.GetLastErrorCode().ToString();
                    }
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("T1: Could not create SelfWithHolding Tax. Self WithHolding was not calculated." + er.Message);
                objResult.Message = "T1: Could not create SelfWithHolding Tax. Self WithHolding was not calculated." + er.Message;
                objResult.MessageCode = "-1";
            }

            return objResult;
        }

        static public void AddRegisterAutoWitholdingTax(SAPbobsCOM.Documents objDoc, string journal, int typoDoc)
        {
            var strSQL = string.Format(Queries.Instance.Queries().Get("GetThridRelated"), objDoc.CardCode);
            var objRS = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            objRS.DoQuery(strSQL);

            var objJournal = (JournalEntries)MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
            objJournal.GetByKey(int.Parse(journal));

            var objCardCode = (BusinessPartners)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
            objCardCode.GetByKey(objDoc.CardCode);

            var lSelfWithHolding = getSelfWithholdingTax(objDoc, objDoc.DocObjectCodeEx);

            var cs = MainObject.Instance.B1Company.GetCompanyService();
            var gs = cs.GetGeneralService("HCO_FWT1100");
            var gd = (GeneralData)gs.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);
                gd.SetProperty("U_DocType", typoDoc);
                gd.SetProperty("U_RlPyCode", objRS.Fields.Item("Code").Value.ToString());
                gd.SetProperty("U_DocDate", DateTime.Now);
                gd.SetProperty("U_RlPyName", objRS.Fields.Item("Name").Value.ToString());
                gd.SetProperty("U_CardCode", objDoc.CardCode);
                gd.SetProperty("U_DocEntry", objDoc.DocEntry);
                gd.SetProperty("U_DocNum", objDoc.DocNum);
                gd.SetProperty("U_TransId", int.Parse(journal));
                gd.SetProperty("U_DocTotal", objDoc.DocTotal);
                gd.SetProperty("U_Comments", "Autoretencion generada automaticamente");
                gd.SetProperty("U_OpType", 1);
                gd.SetProperty("U_TipReg", "2");

            var gdc = gd.Child("HCO_WT1101");

            for (int i = 0; i < lSelfWithHolding.Count; i++)
            {
                var gchc = gdc.Add();
                gchc.SetProperty("U_WTType", lSelfWithHolding[i].TypeLine);
                gchc.SetProperty("U_WTCode", lSelfWithHolding[i].Code);
                gchc.SetProperty("U_WTRate", lSelfWithHolding[i].Percentage);
                gchc.SetProperty("U_WTBase", lSelfWithHolding[i].dbBaseAmount);
                gchc.SetProperty("U_WTAmnt", lSelfWithHolding[i].dbWtAmount);
                gchc.SetProperty("U_OpType", 1);
                gchc.SetProperty("U_BaseLine", "");
                gchc.SetProperty("U_AccountAW", lSelfWithHolding[i].Credit); 
                gchc.SetProperty("U_Account", lSelfWithHolding[i].Debit);
                gchc.SetProperty("U_WTDocAmnt", "0");
            }

            gs.Add(gd);
        }

        private static List<SelfWithholdingTaxInfo> getSelfWithholdingTax(SAPbobsCOM.Documents objDoc, string DocType)
        {
            List<SelfWithholdingTaxInfo> lSWTH = new List<SelfWithholdingTaxInfo>();
            string strSQL = "";
            try
            {
                RelatedParties.Instance.GetRelPartyConfiguration();
                string typeFilter = "'S'";
                var strSQLApply = Queries.Instance.Queries().Get("ApplyAutoWitholdingTax");
                var objRSApplt = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objRSApplt.DoQuery(strSQLApply);

                if (objRSApplt.RecordCount > 0)
                {
                    if (objRSApplt.Fields.Item("U_SWTLiable").Value.ToString().Equals("Y"))
                        typeFilter += ", 'G'";
                }

                strSQL = string.Format(Queries.Instance.Queries().Get(DocType.Equals("13") ? "GetSelfWithHoldingTax" : "GetSelfWithHoldingTaxNC"), objDoc.DocEntry, typeFilter);

                SAPbobsCOM.Recordset objRS = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objRS.DoQuery(strSQL);
                while (!objRS.EoF)
                {
                    if (double.Parse(objRS.Fields.Item("U_Rate").Value.ToString()) > 0)
                    {
                        SelfWithholdingTaxInfo objSWI = new SelfWithholdingTaxInfo();
                        objSWI.Code = objRS.Fields.Item("Code").Value.ToString();

                        if (DocType.Equals("13"))
                        {
                            objSWI.Credit = objRS.Fields.Item("U_CreditAcct").Value.ToString();
                            objSWI.Debit = objRS.Fields.Item("U_DebitAcct").Value.ToString();
                        }
                        else
                        {
                            objSWI.Debit = objRS.Fields.Item("U_CreditAcct").Value.ToString();
                            objSWI.Credit = objRS.Fields.Item("U_DebitAcct").Value.ToString();
                        }

                        objSWI.TypeLine = objRS.Fields.Item("U_Type").Value.ToString().Equals("G") ? "7" : "8";
                        objSWI.Percentage = double.Parse(objRS.Fields.Item("U_Rate").Value.ToString());
                        objSWI.dbBaseAmount = double.Parse(objRS.Fields.Item("Base").Value.ToString());
                        objSWI.dbWtAmount = double.Parse(objRS.Fields.Item("Retencion").Value.ToString());
                        lSWTH.Add(objSWI);
                    }
                    objRS.MoveNext();
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                lSWTH = null;
            }

            return lSWTH;
        }

        #endregion

        #region add Missing Self Withholding Tax

        public static void loadMissingSWTaxForm()
        {
            try
            {
                //SAPbouiCOM.Form objForm = T1.B1.Base.UIOperations.Operations.openFormfromXML(T1.B1.WithholdingTax.SWTaxResources.AutoretencionesFaltantes, Settings._SelfWithHoldingTax.MissingSWTFormUID, false);
                //objForm.VisibleEx = true;
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void getMissingSWTaxDocuments(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Grid specific = null;
            SAPbouiCOM.DataTable objDTDocuments = null;
            SAPbouiCOM.GridColumn objGridDocuments = null;
            SAPbouiCOM.UserDataSource startDate = null;
            SAPbouiCOM.UserDataSource endDate = null;
            SAPbouiCOM.UserDataSource salesWtax = null;
            SAPbouiCOM.UserDataSource purchWtax = null;
            string strCHKPurch = "N";
            string strCHKSales = "N";

            string str = "";
            string str1 = "";
            string str2 = "";
            try
            {
                SAPbouiCOM.Form objForm = MainObject.Instance.B1Application.Forms.Item(FormUID);
                startDate = objForm.DataSources.UserDataSources.Item("getInDate");
                str1 = startDate.ValueEx.Trim();
                if (str1.Trim().Length == 0)
                {
                    MainObject.Instance.B1Application.MessageBox("Por favor seleccione la fecha de inicio.", 1, "Ok", "", "");
                }
                else
                {
                    endDate = objForm.DataSources.UserDataSources.Item("getEndDate");
                    str2 = endDate.ValueEx.Trim();
                    if (str2.Trim().Length == 0)
                    {
                        MainObject.Instance.B1Application.MessageBox("Por favor seleccione la fecha de fin.", 1, "Ok", "", "");
                    }
                    else
                    {


                        salesWtax = objForm.DataSources.UserDataSources.Item("udsSales");
                        strCHKSales = salesWtax.ValueEx.Trim() == "" ? "N" : salesWtax.ValueEx.Trim();
                        purchWtax = objForm.DataSources.UserDataSources.Item("udsPurch");
                        strCHKPurch = purchWtax.ValueEx.Trim() == "" ? "N" : purchWtax.ValueEx.Trim();



                        if (strCHKSales == "N" && strCHKPurch == "N")
                        {
                            MainObject.Instance.B1Application.MessageBox("Por favor seleccione el tipo de autoretención que desea buscar.", 1, "Ok", "", "");
                        }
                        else
                        {
                            string str3 = Settings._SelfWithHoldingTax.getMissingSWT;

                            str3 = str3.Replace("[--StartDate--]", str1);
                            str3 = str3.Replace("[--EndDate--]", str2);
                            objDTDocuments = objForm.DataSources.DataTables.Item("dtSelfWT");
                            objDTDocuments.ExecuteQuery(str3);
                            if (objDTDocuments.Rows.Count <= 0)
                            {
                                MainObject.Instance.B1Application.MessageBox("No se encontraron documentos faltantes en la fecha especificada", 1, "Ok", "", "");
                            }
                            else
                            {
                                specific = (Grid)objForm.Items.Item("grdSWT").Specific;

                                objGridDocuments = specific.Columns.Item(0);
                                objGridDocuments.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                                objGridDocuments.Editable = true;

                                objGridDocuments = specific.Columns.Item(1);
                                SAPbouiCOM.EditTextColumn oCol = (SAPbouiCOM.EditTextColumn)objGridDocuments;
                                oCol.LinkedObjectType = "13";
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = false;



                                objGridDocuments = specific.Columns.Item(2);
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = false;

                                objGridDocuments = specific.Columns.Item(3);
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = false;

                                objGridDocuments = specific.Columns.Item(4);
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = true;

                                objGridDocuments = specific.Columns.Item(5);
                                objGridDocuments.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = false;

                                specific.AutoResizeColumns();
                                objForm.Items.Item("grdSWT").Visible = true;
                                objForm.Items.Item("btnCalc").Visible = true;
                            }
                        }
                    }
                }
            }
            catch (COMException cOMException1)
            {
                COMException cOMException = cOMException1;
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", cOMException.ErrorCode, "::", cOMException.Message, "::", cOMException.StackTrace })));
                _Logger.Error("", exception);
            }
            catch (Exception exception2)
            {
                Exception exception1 = exception2;
                _Logger.Error("", exception1);

            }
        }

        public static void addMisingSWTDocuments(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.DataTable variable = null;
            SAPbouiCOM.Form oForm = null;
            SAPbobsCOM.Documents objDoc = null;

            XmlDocument xmlDocument = new XmlDocument();
            XmlNodeList xmlNodeLists = null;

            try
            {
                oForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                oForm.Freeze(true);
                variable = oForm.DataSources.DataTables.Item("dtSelfWT");
                xmlDocument.LoadXml(variable.SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly));

                xmlNodeLists = xmlDocument.SelectNodes("/DataTable/Rows/Row[./Cells/Cell[1]/Value/text() = 'Y']");

                Dictionary<int, List<string>> objResult = new Dictionary<int, List<string>>();


                int countJE = xmlNodeLists.Count;
                int intProgress = 1;
                T1.B1.Base.UIOperations.Operations.startProgressBar("Iniciando registro", countJE * 2);
                foreach (XmlNode xmlNodes in xmlNodeLists)
                {

                    string innerText = xmlNodes.SelectSingleNode("Cells/Cell[2]/Value").InnerText;

                    objDoc = (Documents)MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);

                    List<string> ListMessage = new List<string>();

                    if (objDoc.GetByKey(Convert.ToInt32(innerText)))
                    {
                        Base.UIOperations.Operations.setProgressBarMessage("Calculando el documento " + innerText, intProgress);
                        SelfWothholdingTaxResult objResultOperation = calcSelfWTax(objDoc, "13");
                        ListMessage.Add(objResultOperation.MessageCode + " " + objResultOperation.Message);

                    }
                    else
                    {
                        _Logger.Error("No se pudo recuperar el documento " + innerText);
                    }
                    objResult.Add(Convert.ToInt32(innerText), ListMessage);

                    intProgress++;
                }

                int intLastLine = -1;
                for (int i = 0; i < variable.Rows.Count; i++)
                {
                    Base.UIOperations.Operations.setProgressBarMessage("Actualizando resultados", intProgress);
                    int strJournalEntry = int.Parse(variable.GetValue(1, i).ToString());
                    if (objResult.ContainsKey(strJournalEntry))
                    {
                        List<string> objListMsg = objResult[strJournalEntry];
                        string strMsg = "";
                        for (int k = 0; k < objListMsg.Count; k++)
                        {
                            strMsg += objListMsg[k] + ".";
                        }
                        variable.SetValue(5, i, strMsg);
                        intLastLine = i;
                        intProgress++;
                    }
                }


                //SAPbouiCOM.Grid specific = oForm.Items.Item("grdSWT").Specific;
                //SAPbouiCOM.GridColumn objGridDocuments = specific.Columns.Item(0);
                //objGridDocuments.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                //objGridDocuments.Editable = false;
                Base.UIOperations.Operations.stopProgressBar();

                MainObject.Instance.B1Application.MessageBox("La operación finalizó con éxito. Por favor revise los resultados en el listado", 1, "Ok", "", "");


            }
            catch (COMException cOMException1)
            {
                COMException cOMException = cOMException1;
                Exception exception2 = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", cOMException.ErrorCode, "::", cOMException.Message, "::", cOMException.StackTrace })));
                _Logger.Error("", exception2);
                Base.UIOperations.Operations.stopProgressBar();
            }
            catch (Exception exception4)
            {
                Exception exception3 = exception4;
                _Logger.Error("", exception3);
                Base.UIOperations.Operations.stopProgressBar();
            }
            finally
            {
                if (oForm != null)
                {
                    oForm.Freeze(false);
                }
            }
        }


        #endregion

        #region cancelSWtax Wizard


        public static void setSelectedCode(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.DataTable objDT = null;
            SAPbouiCOM.UserDataSource oUDS = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.ChooseFromListEvent oCFLE = null;
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                if (objForm != null)
                {
                    oCFLE = (SAPbouiCOM.ChooseFromListEvent)pVal;
                    objDT = oCFLE.SelectedObjects;
                    if (!objDT.IsEmpty)
                    {
                        oUDS = objForm.DataSources.UserDataSources.Item("UD_SWTC");
                        oUDS.ValueEx = objDT.GetValue("Code", 0).ToString();
                    }

                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }
        private static bool cancelSWtaxPosting(int JournalEntry, out string Result)
        {
            bool blResult = false;
            Result = "";
            try
            {
                SAPbobsCOM.JournalEntries objJE = (JournalEntries)MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                objJE.GetByKey(JournalEntry);
                if (objJE.Cancel() != 0)
                {
                    Result = "No se pudo cancelar al asiento de Autoretención " + JournalEntry.ToString() + "." + MainObject.Instance.B1Company.GetLastErrorDescription();
                    _Logger.Error("Could not cancel JE " + JournalEntry.ToString() + "." + MainObject.Instance.B1Company.GetLastErrorDescription());

                }
                else
                {
                    blResult = true;
                    Result = "El asiento " + JournalEntry.ToString() + " se canceló con éxito";
                }


            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            return blResult;
        }
        private static bool cancelSWTaxRegistration(int JournalEntry, out string Result)
        {
            bool blResult = false;
            Result = "";
            try
            {
                SAPbobsCOM.CompanyService companyService = null;
                SAPbobsCOM.GeneralService generalService = null;
                SAPbobsCOM.GeneralData generalData = null;
                SAPbobsCOM.GeneralDataParams generalDataParams = null;
                companyService = MainObject.Instance.B1Company.GetCompanyService();
                generalService = companyService.GetGeneralService(Settings._SelfWithHoldingTax.SWtaxUDOTransaction);

                SAPbobsCOM.Recordset objRs = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objRs.DoQuery(Settings._SelfWithHoldingTax.getRegistrationFromJEQuery.Replace("[--JE--]", JournalEntry.ToString()));
                while (!objRs.EoF)
                {
                    int DocEntry = int.Parse(objRs.Fields.Item("DocEntry").Value.ToString());

                    generalDataParams = (GeneralDataParams)generalService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    generalDataParams.SetProperty("DocEntry", DocEntry);


                    try
                    {
                        //generalData = generalService.GetByParams(generalDataParams);
                        //generalData.SetProperty("Canceled", "Y");
                        //generalService.Update(generalData);
                        //generalData.

                        generalService.Cancel(generalDataParams);
                        generalData = generalService.GetByParams(generalDataParams);
                        string strResult = generalData.GetProperty("Canceled").ToString();
                        Result += "Internal registration " + DocEntry.ToString() + " canceled";
                    }
                    catch (Exception er)
                    {
                        _Logger.Error("Could not cancel internal registration " + DocEntry.ToString() + "." + er.Message);
                        Result += "Could not cancel internal registration " + DocEntry.ToString() + "." + er.Message;
                    }

                    objRs.MoveNext();
                }

                if (Result.Length == 0)
                {
                    Result = "No Internal registration found";
                }



            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }

            return blResult;
        }
        public static void loadCancelSWTaxForm()
        {
            SAPbouiCOM.Item objItem = null;
            try
            {
                //SAPbouiCOM.Form objForm = T1.B1.Base.UIOperations.Operations.openFormfromXML(T1.B1.WithholdingTax.SWTaxResources.CancelarAutoretenciones, Settings._SelfWithHoldingTax.CancelFormUID, false);
                //if (objForm != null)
                //{

                //    if (!Settings._SelfWithHoldingTax.TransactionCodeBase)
                //    {
                //        objItem = objForm.Items.Item("Item_6");
                //        objItem.Visible = true;
                //        objItem = objForm.Items.Item("txtSWTCode");
                //        objItem.Visible = true;
                //    }
                //    objForm.VisibleEx = true;
                //}
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }


        public static void getPostedSWTaxDocuments(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Grid specific = null;
            SAPbouiCOM.DataTable objDTDocuments = null;
            SAPbouiCOM.GridColumn objGridDocuments = null;
            SAPbouiCOM.UserDataSource startDate = null;
            SAPbouiCOM.UserDataSource endDate = null;
            SAPbouiCOM.UserDataSource sWtaxCode = null;
            string str = "";
            string str1 = "";
            string str2 = "";
            try
            {
                SAPbouiCOM.Form objForm = MainObject.Instance.B1Application.Forms.Item(FormUID);
                startDate = objForm.DataSources.UserDataSources.Item("getInDate");
                str1 = startDate.ValueEx.Trim();
                if (str1.Trim().Length == 0)
                {
                    MainObject.Instance.B1Application.MessageBox("Por favor seleccione la fecha de inicio.", 1, "Ok", "", "");
                }
                else
                {
                    endDate = objForm.DataSources.UserDataSources.Item("getEndDate");
                    str2 = endDate.ValueEx.Trim();
                    if (str2.Trim().Length == 0)
                    {
                        MainObject.Instance.B1Application.MessageBox("Por favor seleccione la fecha de fin.", 1, "Ok", "", "");
                    }
                    else
                    {
                        sWtaxCode = objForm.DataSources.UserDataSources.Item("UD_SWTC");
                        str = sWtaxCode.ValueEx.Trim();
                        if (Settings._SelfWithHoldingTax.TransactionCodeBase)
                        {
                            str = "All";
                        }


                        if (str.Trim().Length == 0)
                        {
                            MainObject.Instance.B1Application.MessageBox("Por favor seleccione el codigo de autoretención.", 1, "Ok", "", "");
                        }
                        else
                        {
                            string str3 = Settings._SelfWithHoldingTax.getPostedSWtaxQueryV1;
                            if (Settings._SelfWithHoldingTax.TransactionCodeBase)
                            {
                                str3 = Settings._SelfWithHoldingTax.getPostedSWtaxQueryV2.Replace("[--TransCode--]", Settings._SelfWithHoldingTax.WTaxTransCode);
                            }
                            else
                            {
                                str3 = Settings._SelfWithHoldingTax.getPostedSWtaxQueryV1.Replace("[--SWTCode--]", str);
                            }

                            str3 = str3.Replace("[--StartDate--]", str1);
                            str3 = str3.Replace("[--EndDate--]", str2);
                            objDTDocuments = objForm.DataSources.DataTables.Item("dtSelfWT");
                            objDTDocuments.ExecuteQuery(str3);
                            if (objDTDocuments.Rows.Count <= 0)
                            {
                                MainObject.Instance.B1Application.MessageBox("No se encontraron documentos contabilizados en la fecha especificada", 1, "Ok", "", "");
                            }
                            else
                            {
                                specific = (Grid)objForm.Items.Item("grdSWT").Specific;
                                objGridDocuments = specific.Columns.Item(0);
                                objGridDocuments.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                                objGridDocuments.Editable = true;
                                objGridDocuments = specific.Columns.Item(1);
                                SAPbouiCOM.EditTextColumn oCol = (SAPbouiCOM.EditTextColumn)objGridDocuments;
                                oCol.LinkedObjectType = "30";
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = false;
                                objGridDocuments = specific.Columns.Item(2);
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = false;
                                objGridDocuments = specific.Columns.Item(3);
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = false;
                                objGridDocuments = specific.Columns.Item(4);
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = false;
                                objGridDocuments = specific.Columns.Item(5);
                                objGridDocuments.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = false;
                                specific.AutoResizeColumns();
                                objForm.Items.Item("grdSWT").Visible = true;
                                objForm.Items.Item("btnCalc").Visible = true;
                            }
                        }
                    }
                }
            }
            catch (COMException cOMException1)
            {
                COMException cOMException = cOMException1;
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", cOMException.ErrorCode, "::", cOMException.Message, "::", cOMException.StackTrace })));
                _Logger.Error("", exception);
            }
            catch (Exception exception2)
            {
                Exception exception1 = exception2;
                _Logger.Error("", exception1);

            }
        }

        public static void cancelPostedTaxDocuments(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.DataTable variable = null;
            SAPbouiCOM.Form oForm = null;

            XmlDocument xmlDocument = new XmlDocument();
            XmlNodeList xmlNodeLists = null;

            try
            {
                oForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                oForm.Freeze(true);
                variable = oForm.DataSources.DataTables.Item("dtSelfWT");
                xmlDocument.LoadXml(variable.SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly));

                xmlNodeLists = xmlDocument.SelectNodes("/DataTable/Rows/Row[./Cells/Cell[1]/Value/text() = 'Y']");

                Dictionary<int, List<string>> objResult = new Dictionary<int, List<string>>();
                SAPbouiCOM.DBDataSource variable4 = oForm.DataSources.DBDataSources.Item("@HCO_SW0100");

                int countJE = xmlNodeLists.Count;
                int intProgress = 1;
                T1.B1.Base.UIOperations.Operations.startProgressBar("Iniciando reversión", countJE * 2);
                foreach (XmlNode xmlNodes in xmlNodeLists)
                {

                    string innerText = xmlNodes.SelectSingleNode("Cells/Cell[2]/Value").InnerText;
                    Base.UIOperations.Operations.setProgressBarMessage("Reversando JE " + innerText, intProgress);
                    string strResultJe = "";
                    string strResultInternal = "";
                    List<string> ListMessage = new List<string>();

                    cancelSWtaxPosting(Convert.ToInt32(innerText), out strResultJe);
                    ListMessage.Add(strResultJe);
                    cancelSWTaxRegistration(Convert.ToInt32(innerText), out strResultInternal);
                    ListMessage.Add(strResultInternal);
                    objResult.Add(Convert.ToInt32(innerText), ListMessage);
                    intProgress++;
                }

                int intLastLine = -1;
                for (int i = 0; i < variable.Rows.Count; i++)
                {
                    Base.UIOperations.Operations.setProgressBarMessage("Actualizando resultados", intProgress);
                    int strJournalEntry = int.Parse(variable.GetValue(1, i).ToString());
                    if (objResult.ContainsKey(strJournalEntry))
                    {
                        List<string> objListMsg = objResult[strJournalEntry];
                        string strMsg = "";
                        for (int k = 0; k < objListMsg.Count; k++)
                        {
                            strMsg += objListMsg[k] + ".";
                        }
                        variable.SetValue(5, i, strMsg);
                        intLastLine = i;
                        intProgress++;
                    }
                }

                oForm.Items.Item("Item_7").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                SAPbouiCOM.Grid specific = (Grid)oForm.Items.Item("grdSWT").Specific;
                SAPbouiCOM.GridColumn objGridDocuments = specific.Columns.Item(0);
                objGridDocuments.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                objGridDocuments.Editable = false;
                Base.UIOperations.Operations.stopProgressBar();

                MainObject.Instance.B1Application.MessageBox("La operación finalizó con éxito. Por favor revise los resultados en el listado", 1, "Ok", "", "");


            }
            catch (COMException cOMException1)
            {
                COMException cOMException = cOMException1;
                Exception exception2 = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", cOMException.ErrorCode, "::", cOMException.Message, "::", cOMException.StackTrace })));
                _Logger.Error("", exception2);
                Base.UIOperations.Operations.stopProgressBar();
            }
            catch (Exception exception4)
            {
                Exception exception3 = exception4;
                _Logger.Error("", exception3);
                Base.UIOperations.Operations.stopProgressBar();
            }
            finally
            {
                if (oForm != null)
                {
                    oForm.Freeze(false);
                }
            }
        }


        #endregion

        #region Autoretención Configuración

        public static void loadSWTaxConfigForm()
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.FormCreationParams objParams = null;
            try
            {
                objParams = (FormCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                objParams.XmlData = GetXmlUDO(Settings.SelfWithHoldingTaxTypes.SELFWHITHHOLDINGTAX_CONFIGURAION);
                objParams.FormType = "HCO_FSW0100";
                objParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                objForm = MainObject.Instance.B1Application.Forms.AddEx(objParams);

                InitConfiguration(objForm);

                objForm.VisibleEx = true;

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void InitConfiguration(Form form)
        {
            SAPbouiCOM.ChooseFromList objChooseFromList = form.ChooseFromLists.Item("cflWtCode");

            try
            {
                Conditions conditions = objChooseFromList.GetConditions();
                Condition condition = conditions.Add();
                condition.BracketOpenNum = 1;
                condition.Alias = "U_HCO_WTType";
                condition.Operation = BoConditionOperation.co_EQUAL;
                condition.CondVal = "2";
                condition.Relationship = BoConditionRelationship.cr_AND;
                condition = conditions.Add();
                condition.Alias = "U_HCO_Area";
                condition.Operation = BoConditionOperation.co_EQUAL;
                condition.CondVal = "V";
                condition.BracketCloseNum = 1;
                objChooseFromList.SetConditions(conditions);

                objChooseFromList = form.ChooseFromLists.Item("cflAcctDebit");
                conditions = objChooseFromList.GetConditions();
                condition = conditions.Add();
                condition.BracketOpenNum = 1;
                condition.Alias = "Postable";
                condition.Operation = BoConditionOperation.co_EQUAL;
                condition.CondVal = "Y";
                condition.Relationship = BoConditionRelationship.cr_AND;
                condition = conditions.Add();
                condition.Alias = "LocManTran";
                condition.Operation = BoConditionOperation.co_EQUAL;
                condition.CondVal = "N";
                condition.BracketCloseNum = 1;
                objChooseFromList.SetConditions(conditions);

                objChooseFromList = form.ChooseFromLists.Item("cflAcctCredit");
                conditions = objChooseFromList.GetConditions();
                condition = conditions.Add();
                condition.BracketOpenNum = 1;
                condition.Alias = "Postable";
                condition.Operation = BoConditionOperation.co_EQUAL;
                condition.CondVal = "Y";
                condition.Relationship = BoConditionRelationship.cr_AND;
                condition = conditions.Add();
                condition.Alias = "LocManTran";
                condition.Operation = BoConditionOperation.co_EQUAL;
                condition.CondVal = "N";
                condition.BracketCloseNum = 1;
                objChooseFromList.SetConditions(conditions);

                objChooseFromList = form.ChooseFromLists.Item("cflCardCode");
                conditions = objChooseFromList.GetConditions();
                condition = conditions.Add();
                condition.Alias = "CardType";
                condition.Operation = BoConditionOperation.co_EQUAL;
                condition.CondVal = "C";
                objChooseFromList.SetConditions(conditions);

                AddGroupsToMatrix(form);
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objChooseFromList);
                objChooseFromList = null;
                GC.Collect();
            }
        }

        public static void AddGroupsToMatrix(Form form)
        {
            var recordSet = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                var queryGroup = Queries.Instance.Queries().Get("GetGroupItems");
                Matrix matrix = (Matrix)form.Items.Item("Item_19").Specific;
                recordSet.DoQuery(queryGroup);
                matrix.Clear();
                matrix.AddRow(recordSet.RecordCount);
                matrix.FlushToDataSource();
                for (int i = 0; i < recordSet.RecordCount; i++)
                {
                    form.DataSources.DBDataSources.Item("@HCO_SW0101").SetValue("LineId", i, (i + 1).ToString());
                    form.DataSources.DBDataSources.Item("@HCO_SW0101").SetValue("U_ItmsGrpCod", i, recordSet.Fields.Item("ItmsGrpCod").Value.ToString());
                    form.DataSources.DBDataSources.Item("@HCO_SW0101").SetValue("U_ItmsGrpNam", i, recordSet.Fields.Item("ItmsGrpNam").Value.ToString());
                    form.DataSources.DBDataSources.Item("@HCO_SW0101").SetValue("U_Exclude", i, "N");
                    recordSet.MoveNext();
                }

                matrix.LoadFromDataSourceEx();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(recordSet);
                recordSet = null;
                GC.Collect();
            }

        }

        public static void CheckGroups(Form form)
        {
            var recordSet = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                var queryGroup = Queries.Instance.Queries().Get("GetMissingGroupsFromAutorret");
                recordSet.DoQuery(string.Format(queryGroup, form.DataSources.DBDataSources.Item("@HCO_SW0100").GetValue("Code", 0).ToString()));
                if (recordSet.RecordCount > 0)
                {
                    Matrix matrix = (Matrix)form.Items.Item("Item_19").Specific;
                    int row = matrix.RowCount;
                    matrix.AddRow(recordSet.RecordCount);
                    matrix.FlushToDataSource();
                    while (!recordSet.EoF)
                    {
                        form.DataSources.DBDataSources.Item("@HCO_SW0101").SetValue("LineId", row, (row + 1).ToString());
                        form.DataSources.DBDataSources.Item("@HCO_SW0101").SetValue("U_ItmsGrpCod", row, recordSet.Fields.Item("ItmsGrpCod").Value.ToString());
                        form.DataSources.DBDataSources.Item("@HCO_SW0101").SetValue("U_ItmsGrpNam", row, recordSet.Fields.Item("ItmsGrpNam").Value.ToString());
                        form.DataSources.DBDataSources.Item("@HCO_SW0101").SetValue("U_Exclude", row, "N");
                        recordSet.MoveNext();
                        row++;
                    }

                    matrix.LoadFromDataSourceEx();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(recordSet);
                recordSet = null;
                GC.Collect();
            }

        }

        static public void addInsertRowRelationMenuUDO(SAPbouiCOM.Form objForm, SAPbouiCOM.ContextMenuInfo eventInfo)
        {
            SAPbouiCOM.MenuCreationParams objParams = null;

            try
            {
                objParams = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objParams.String = "Agregar línea";
                objParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                objParams.UniqueID = "HCO_MWTRU";
                objParams.Position = -1;
                objForm.Menu.AddEx(objParams);
                EventInfoClass objEvent = new EventInfoClass();
                objEvent.ColUID = eventInfo.ColUID;
                objEvent.FormUID = eventInfo.FormUID;
                objEvent.ItemUID = eventInfo.ItemUID;
                objEvent.Row = eventInfo.Row;
                CacheManager.CacheManager.Instance.addToCache(Settings._Main.lastRightClickEventInfo, objEvent, CacheManager.CacheManager.objCachePriority.Default);
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

        static public void removeInsertRowRelationMenuUDO()
        {


            try
            {
                if (MainObject.Instance.B1Application.Menus.Exists("HCO_MWTRU"))
                {
                    MainObject.Instance.B1Application.Menus.RemoveEx("HCO_MWTRU");
                }
                CacheManager.CacheManager.Instance.removeFromCache(Settings._Main.lastRightClickEventInfo);


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

        static public void addDeleteRowRelationMenuUDO(SAPbouiCOM.Form objForm, SAPbouiCOM.ContextMenuInfo eventInfo)
        {
            SAPbouiCOM.MenuCreationParams objParams = null;

            try
            {
                objParams = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objParams.String = "Eliminar línea";
                objParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                objParams.UniqueID = "HCO_MWTDRU";
                objParams.Position = -1;
                objForm.Menu.AddEx(objParams);
                EventInfoClass objEvent = new EventInfoClass();
                objEvent.ColUID = eventInfo.ColUID;
                objEvent.FormUID = eventInfo.FormUID;
                objEvent.ItemUID = eventInfo.ItemUID;
                objEvent.Row = eventInfo.Row;
                CacheManager.CacheManager.Instance.addToCache(Settings._Main.lastRightClickEventInfo, objEvent, CacheManager.CacheManager.objCachePriority.Default);


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

        static public void removeDeleteRowRelationMenuUDO()
        {


            try
            {
                if (MainObject.Instance.B1Application.Menus.Exists("HCO_MWTDRU"))
                {
                    MainObject.Instance.B1Application.Menus.RemoveEx("HCO_MWTDRU");
                }
                CacheManager.CacheManager.Instance.removeFromCache(Settings._Main.lastRightClickEventInfo);
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

        static public void MatrixOperationUDO(EventInfoClass eventInfo, string Action)
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.Matrix objMatrix = null;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(eventInfo.FormUID);

                objMatrix = (Matrix)objForm.Items.Item(eventInfo.ItemUID).Specific;

                int intRow = objMatrix.RowCount;
                switch (Action)
                {
                    case "Add":
                        objMatrix.AddRow(1, 1);

                        objMatrix.FlushToDataSource();
                        objForm.DataSources.DBDataSources.Item(eventInfo.ItemUID == "Item_31" ? "@HCO_SW0102" : "@HCO_SW0103").SetValue("LineId", intRow, (intRow + 1).ToString());
                        objForm.DataSources.DBDataSources.Item(eventInfo.ItemUID == "Item_31" ? "@HCO_SW0102" : "@HCO_SW0103").SetValue(eventInfo.ItemUID == "Item_31" ? "U_ItemCode" : "U_CardCode", intRow, "");
                        objForm.DataSources.DBDataSources.Item(eventInfo.ItemUID == "Item_31" ? "@HCO_SW0102" : "@HCO_SW0103").SetValue(eventInfo.ItemUID == "Item_31" ? "U_ItemName" : "U_CardName", intRow, "");

                        objMatrix.LoadFromDataSourceEx();

                        objMatrix.SetCellFocus(intRow + 1, 1);


                        break;
                    case "Delete":
                        objMatrix.DeleteRow(intRow);
                        objMatrix.FlushToDataSource();
                        break;

                }
                if (objForm.Mode == BoFormMode.fm_OK_MODE)
                    objForm.Mode = BoFormMode.fm_UPDATE_MODE;
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

        static public void addAllPBS(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Form objForm = null;
            string strIsSales = "";
            string strIsPurchase = "";
            SAPbouiCOM.DBDataSource objDS = null;
            SAPbouiCOM.Matrix objMatrix = null;
            SAPbobsCOM.Recordset objRS = null;

            SAPbobsCOM.SBObob objBridge = null;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);

                objDS = objForm.DataSources.DBDataSources.Item("@HCO_SW0100");
                strIsSales = objDS.GetValue("U_Sales", objDS.Offset);
                strIsPurchase = objDS.GetValue("U_Purchase", objDS.Offset);
                objMatrix = (Matrix)objForm.Items.Item("0_U_G").Specific;
                objMatrix.Clear();
                objForm.Freeze(true);
                objMatrix.FlushToDataSource();
                objBridge = (SBObob)MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                int intRow = 0;
                objForm.Freeze(true);

                T1.B1.Base.UIOperations.Operations.startProgressBar("Agregando socios de negocio...", 2);
                if (strIsSales == "Y")
                {

                    objRS = objBridge.GetBPList(SAPbobsCOM.BoCardTypes.cCustomer);
                    while (!objRS.EoF)
                    {
                        string strCardCode = objRS.Fields.Item("CardCode").Value.ToString();
                        string strCardName = objRS.Fields.Item("CardName").Value.ToString();

                        objMatrix.AddRow(1, intRow);

                        objMatrix.SetCellWithoutValidation(intRow + 1, "C_0_1", strCardCode);
                        objMatrix.SetCellWithoutValidation(intRow + 1, "C_0_2", strCardName);
                        objMatrix.FlushToDataSource();
                        intRow++;
                        objRS.MoveNext();
                    }
                }
                if (strIsPurchase == "Y")
                {

                    objRS = objBridge.GetBPList(SAPbobsCOM.BoCardTypes.cSupplier);
                    while (!objRS.EoF)
                    {
                        string strCardCode = objRS.Fields.Item("CardCode").Value.ToString();
                        string strCardName = objRS.Fields.Item("CardName").Value.ToString();

                        objMatrix.AddRow(1, intRow);

                        objMatrix.SetCellWithoutValidation(intRow + 1, "C_0_1", strCardCode);
                        objMatrix.SetCellWithoutValidation(intRow + 1, "C_0_2", strCardName);
                        objMatrix.FlushToDataSource();
                        intRow++;
                        objRS.MoveNext();
                    }
                }
                objForm.Freeze(false);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                if (objForm != null)
                {
                    objForm.Freeze(false);
                }
                T1.B1.Base.UIOperations.Operations.stopProgressBar();
                T1.B1.Base.UIOperations.Operations.setStatusBarMessage("Operación finalizada.", false, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
        }

        static public void clearAllPBS(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Form objForm = null;

            SAPbouiCOM.Matrix objMatrix = null;


            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);

                objMatrix = (Matrix)objForm.Items.Item("0_U_G").Specific;
                objMatrix.Clear();

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                if (objForm != null)
                {
                    objForm.Freeze(false);
                }
            }
        }

        static public void SelectFieldCFL(SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                var form = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                var field = "";
                string code = "";
                string name = "";
                string nameFieldDataSource = "";
                string datasource = "";
                bool loadFromDataSource = false;
                if (new string[] { "Item_31", "Item_12" }.Contains(pVal.ItemUID))
                {
                    datasource = ((Matrix)form.Items.Item(pVal.ItemUID).Specific).Columns.Item(pVal.ColUID).DataBind.TableName;
                    field = ((Matrix)form.Items.Item(pVal.ItemUID).Specific).Columns.Item(pVal.ColUID).DataBind.Alias;
                    loadFromDataSource = true;
                }
                else
                {
                    datasource = ((EditText)form.Items.Item(pVal.ItemUID).Specific).DataBind.TableName;
                    field = ((EditText)form.Items.Item(pVal.ItemUID).Specific).DataBind.Alias;
                }
                switch (pVal.ItemUID)
                {
                    case "Item_3":
                        code = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "AcctCode")[0].ToString();
                        name = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "AcctName")[0].ToString();
                        nameFieldDataSource = "U_DebitAcctN";
                        break;
                    case "Item_9":
                        code = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "AcctCode")[0].ToString();
                        name = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "AcctName")[0].ToString();
                        nameFieldDataSource = "U_CreditAcctN";
                        break;
                    case "Item_16":
                        code = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "WTCode")[0].ToString();
                        break;
                    case "Item_31":
                        code = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "ItemCode")[0].ToString();
                        name = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "ItemName")[0].ToString();
                        nameFieldDataSource = "U_ItemName";
                        break;
                    case "Item_12":
                        code = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "CardCode")[0].ToString();
                        name = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "CardName")[0].ToString();
                        nameFieldDataSource = "U_CardName";
                        break;
                    case "Item_34":
                        code = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Name")[0].ToString();
                        name = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Code")[0].ToString();
                        nameFieldDataSource = "U_TaxCardTypeID";
                        break;
                }
                if (string.IsNullOrEmpty(code))
                    return;

                form.DataSources.DBDataSources.Item(datasource).SetValue(field, pVal.Row > 0 ? pVal.Row - 1 : 0, code);
                if (!string.IsNullOrEmpty(nameFieldDataSource))
                    form.DataSources.DBDataSources.Item(datasource).SetValue(nameFieldDataSource, pVal.Row > 0 ? pVal.Row - 1 : 0, name);
                if (loadFromDataSource)
                    ((Matrix)form.Items.Item(pVal.ItemUID).Specific).LoadFromDataSourceEx();

                if (form.Mode == BoFormMode.fm_OK_MODE)
                    form.Mode = BoFormMode.fm_UPDATE_MODE;

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        static public void EnableItems(bool enable, Form objForm)
        {
            try
            {
                objForm.Items.Item("Item_8").Enabled = enable;
                objForm.Items.Item("Item_30").Enabled = enable;
                objForm.Items.Item("Item_30").Enabled = enable;
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        static public bool ValidateFields(ItemEvent pVal)
        {
            bool BubbelEvent = true;
            try
            {
                var form = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                if (form.Mode == BoFormMode.fm_ADD_MODE)
                {
                    if (string.IsNullOrEmpty(form.DataSources.DBDataSources.Item("@HCO_SW0100").GetValue("U_DebitAcct", 0)))
                    {
                        MainObject.Instance.B1Application.StatusBar.SetText("Debe seleccionar una cuenta débito", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        form.Items.Item("Item_3").Click(BoCellClickType.ct_Regular);
                        BubbelEvent = false;
                    }
                    else if (string.IsNullOrEmpty(form.DataSources.DBDataSources.Item("@HCO_SW0100").GetValue("U_CreditAcct", 0)))
                    {
                        MainObject.Instance.B1Application.StatusBar.SetText("Debe seleccionar una cuenta crédito", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        form.Items.Item("Item_9").Click(BoCellClickType.ct_Regular);
                        BubbelEvent = false;
                    }
                    else if (string.IsNullOrEmpty(form.DataSources.DBDataSources.Item("@HCO_SW0100").GetValue("U_Type", 0)))
                    {
                        MainObject.Instance.B1Application.StatusBar.SetText("Debe seleccionar un tipo", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        form.Items.Item("Item_30").Click(BoCellClickType.ct_Regular);
                        BubbelEvent = false;
                    }
                    else if (string.IsNullOrEmpty(form.DataSources.DBDataSources.Item("@HCO_SW0100").GetValue("U_WTCode", 0))
                        && form.DataSources.DBDataSources.Item("@HCO_SW0100").GetValue("U_Type", 0) == "G")
                    {
                        MainObject.Instance.B1Application.StatusBar.SetText("Debe seleccionar una retención vinculada para el tipo general", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        form.Items.Item("Item_16").Click(BoCellClickType.ct_Regular);
                        BubbelEvent = false;
                    }
                    else if (!string.IsNullOrEmpty(form.DataSources.DBDataSources.Item("@HCO_SW0102").GetValue("U_ItemCode", 0))
                        && string.IsNullOrEmpty(form.DataSources.DBDataSources.Item("@HCO_SW0100").GetValue("U_ItemAction", 0)))
                    {
                        MainObject.Instance.B1Application.StatusBar.SetText("Debe seleccionar una acción en la pestaña de articulos", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        form.Items.Item("Item_24").Click(BoCellClickType.ct_Regular);
                        BubbelEvent = false;
                    }
                }
                else if (form.Mode == BoFormMode.fm_ADD_MODE || form.Mode == BoFormMode.fm_UPDATE_MODE)
                {
                    if (string.IsNullOrEmpty(form.DataSources.DBDataSources.Item("@HCO_SW0100").GetValue("U_TaxCardType", 0)))
                        form.DataSources.DBDataSources.Item("@HCO_SW0100").SetValue("U_TaxCardTypeID", 0, "");
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }

            return BubbelEvent;
        }

        #endregion

        #region SelfwithHlding in Documents

        static public void HCOSelfWithHoldingFolderAdd(string strFormUID)
        {

            SAPbouiCOM.Form objForm = null;
            int intLeft = 0;
            string strUID = "";
            SAPbouiCOM.Item objItemBase = null;
            SAPbouiCOM.Item objItem = null;
            //SAPbouiCOM.Matrix objMatrix = null;

            SAPbouiCOM.Folder objFolder = null;
            XmlDocument xmlResult = null;
            SAPbouiCOM.BoFormMode objMode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
            bool blFolderFound = false;
            XmlDocument objFormXML = null;
            XmlNode objNode = null;

            try
            {

                objForm = MainObject.Instance.B1Application.Forms.Item(strFormUID);
                objMode = objForm.Mode;

                objFormXML = new XmlDocument();
                objFormXML.LoadXml(objForm.GetAsXML());
                objNode = objFormXML.SelectSingleNode("/Application/forms/action/form/items/action/item[@uid='HCO_FSWT']");
                if (objNode != null)
                {
                    blFolderFound = true;
                }
                objForm.Freeze(true);
                if (blFolderFound)
                {
                    objForm.Freeze(true);
                    objItem = objForm.Items.Item(Settings._SelfWithHoldingTax.SelfWithHoldingFolderId);
                }
                else
                {
                    objForm.Freeze(true);
                    string strFolderXML = SWTaxResources.SWTaxFolderDocuments;
                    strFolderXML = strFolderXML.Replace("[--UniqueId--]", strFormUID);
                    MainObject.Instance.B1Application.LoadBatchActions(ref strFolderXML);
                    string strResult = MainObject.Instance.B1Application.GetLastBatchResults();
                    xmlResult = new XmlDocument();
                    xmlResult.LoadXml(strResult);
                    bool errors = xmlResult.SelectSingleNode("/result/errors").HasChildNodes != true ? false : true;
                    if (!errors)
                    {
                        objItem = objForm.Items.Item(Settings._SelfWithHoldingTax.SelfWithHoldingFolderId);

                    }
                    else
                    {
                        objItem = null;
                    }
                }

                objForm.Freeze(false);

                if (objItem != null)
                {
                    objForm.Freeze(true);
                    #region Folder

                    objItemBase = objForm.Items.Item(Settings._SelfWithHoldingTax.lastFolderId);
                    if (objItemBase != null)
                    {
                        if (objItemBase.Visible)
                        {
                            intLeft = objItemBase.Left;
                            strUID = objItemBase.UniqueID;
                            objItem.Left = intLeft + 1;
                            objItem.FromPane = 0;
                            objItem.ToPane = 0;
                            objFolder = (Folder)objItem.Specific;
                            objFolder.GroupWith(strUID);
                            objFolder.Item.TextStyle = 1;
                            objItem.Visible = true;
                        }
                    }
                    #endregion Folder;

                    objForm.Items.Item("HCO_IT01").TextStyle = 1;
                    //objForm.Mode = objMode;
                    objForm.Freeze(false);
                }
                else
                {
                    //objForm.Mode = objMode;
                }

                //}



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

            if (objForm != null) objForm.Freeze(false);
        }

        static public void OpenSelfWitholdingRecord(SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                var formDoc = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                var grid = (Grid)formDoc.Items.Item("HCO_GR01").Specific;

                MainObject.Instance.B1Application.Menus.Item("HCO_MWT0003").Activate();
                var formAuto = MainObject.Instance.B1Application.Forms.ActiveForm;
                formAuto.Freeze(true);

                try
                {
                    MainObject.Instance.B1Application.Menus.Item("1281").Activate();
                    ((EditText)formAuto.Items.Item("1_U_E").Specific).Value = grid.DataTable.GetValue(pVal.ColUID, pVal.Row).ToString();
                    formAuto.Items.Item("1").Click();
                }
                finally
                {
                    formAuto.Freeze(false);
                }
            }
            finally
            {

            }
        }

        static public void getSWTaxInfoForDocument(SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            SAPbobsCOM.Documents objDoc = null;
            string strDocType = "";
            string strSQL = "";
            SAPbouiCOM.DataTable objDT = null;
            SAPbouiCOM.Grid objGrid = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.EditTextColumn oEditTExt = null;
            SAPbouiCOM.GridColumn objGridColumn = null;
            try
            {
                strDocType = BusinessObjectInfo.Type;
                objDoc = (Documents)MainObject.Instance.B1Company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), strDocType));
                if (objDoc.Browser.GetByKeys(BusinessObjectInfo.ObjectKey))
                {
                    objForm = MainObject.Instance.B1Application.Forms.Item(BusinessObjectInfo.FormUID);
                    objForm.Freeze(true);
                    try
                    {
                        objDT = objForm.DataSources.DataTables.Item("HCO_DSWT");
                    }
                    catch
                    {
                        //HCOSelfWithHoldingFolderAdd(BusinessObjectInfo.FormUID);
                        objDT = objForm.DataSources.DataTables.Item("HCO_DSWT");
                    }
                    if (objDT != null)
                    {
                        strSQL = string.Format(Queries.Instance.Queries().Get("GetWitholdingRegistered"), strDocType, objDoc.DocEntry.ToString());

                        objDT.ExecuteQuery(strSQL);

                        objGrid = (Grid)objForm.Items.Item("HCO_GR01").Specific;

                        objGridColumn = objGrid.Columns.Item(0);
                        objGridColumn.Editable = false;
                        objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                        oEditTExt = (SAPbouiCOM.EditTextColumn)objGridColumn;
                        oEditTExt.LinkedObjectType = "HCO_FWT1100";

                        objGridColumn = objGrid.Columns.Item(1);
                        objGridColumn.Editable = false;
                        objGridColumn.Visible = false;
                        objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;

                        objGridColumn = objGrid.Columns.Item(2);
                        objGridColumn.Editable = false;
                        objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                        oEditTExt = (SAPbouiCOM.EditTextColumn)objGridColumn;

                        objGridColumn = objGrid.Columns.Item(3);
                        objGridColumn.Editable = false;
                        objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                        oEditTExt = (SAPbouiCOM.EditTextColumn)objGridColumn;

                        objGridColumn = objGrid.Columns.Item(4);
                        objGridColumn.Editable = false;
                        objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;

                        objGridColumn = objGrid.Columns.Item(5);
                        objGridColumn.Editable = false;
                        objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                        oEditTExt = (SAPbouiCOM.EditTextColumn)objGridColumn;
                        oEditTExt.LinkedObjectType = "30";
                        oEditTExt.RightJustified = true;
                    }
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                if (objForm != null)
                {
                    objForm.Freeze(false);
                }
            }
        }
        #endregion

    }
}
