using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using System.Runtime.InteropServices;
using System.Xml;
using SAPbobsCOM;
using System.Globalization;
using System.ComponentModel;
using Newtonsoft.Json;
using System.Drawing;
using SAPbouiCOM;
using System.Data;

namespace T1.B1.IvaCosto
{
    public class IvaCosto
    {
        private static IvaCosto objIvaCosto;
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private IvaCosto()
        {
            if (objIvaCosto == null) objIvaCosto = new IvaCosto();
        }

        #region Carga de formulario de transacción- implementar después
        public static string GetXmlForm(Settings.EIvaCosto type)
        {
            XmlDocument oXML = new XmlDocument();
            switch (type)
            {
                case Settings.EIvaCosto.IVA:
                    oXML.Load(AppDomain.CurrentDomain.BaseDirectory + "\\Forms\\IvaCosto.srf");
                    break;

                default:
                    return string.Empty;
            }
            return oXML.InnerXml;
        }
        public static void LoadWIvaCostoForm(Settings.EIvaCosto type)
        {
            try
            {
                FormCreationParams objParams = (FormCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
                objParams.XmlData = GetXmlForm(type);
                //objParams.FormType = GetTypeUDO(type);
                objParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);

                Form objForm = MainObject.Instance.B1Application.Forms.AddEx(objParams);
                objForm.VisibleEx = true;

            }
            catch (Exception er)
            {
                _Logger.Error("(LoadWIvaCostoForm)", er);
            }

        }
        public static string GetTypeUDO(Settings.EIvaCosto type)
        {
            switch (type)
            {
                case Settings.EIvaCosto.IVA:
                    return "HCO_FWT0100";

                default:
                    return string.Empty;
            }
        }
        #endregion

        #region Worker
        static public void addDocumentInfo(AddDocumentInfoArgs BusinessObjectInfo)
        {
            BackgroundWorker addDocumentInfoWorker = new BackgroundWorker();
            addDocumentInfoWorker.WorkerSupportsCancellation = false;
            addDocumentInfoWorker.WorkerReportsProgress = false;
            addDocumentInfoWorker.DoWork += AddDocumentInfoWorker_DoWork;
            addDocumentInfoWorker.RunWorkerCompleted += AddDocumentInfoWorker_RunWorkerCompleted;
            addDocumentInfoWorker.RunWorkerAsync(BusinessObjectInfo);
        }
        static private void AddDocumentInfoWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!(e.Error == null))
            {
                _Logger.Error(e.Error.Message);

            }

            else
            {
                //this.tbProgress.Text = "Done!";
            }
        }
        static private void AddDocumentInfoWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            AddDocumentInfoArgs oInfo = null;
            SAPbobsCOM.Documents objDoc = null;
            List<string> ICPurchaseDocuments = new List<string>();

            try
            {
                oInfo = (AddDocumentInfoArgs)e.Argument;
                ICPurchaseDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._IvaCosto.ICPurchaseObjects);
                objDoc = (SAPbobsCOM.Documents)MainObject.Instance.B1Company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), oInfo.ObjectType));

                if (ICPurchaseDocuments.Contains(oInfo.FormtTypeEx))
                {
                    objDoc.GetByKey(int.Parse(oInfo.ObjectKey));

                    if (HasIVACosto(objDoc))
                    {
                        switch (objDoc.DocObjectCode)
                        {
                            case BoObjectTypes.oPurchaseDeliveryNotes:
                                CreateRevaluation(true, objDoc);
                                break;
                            case BoObjectTypes.oPurchaseReturns:
                                CreateRevaluation(false, objDoc);
                                break;
                            case BoObjectTypes.oPurchaseInvoices:
                                CheckCapitalizationGenerated(objDoc.DocEntry.ToString(), objDoc);
                                CreateRevaluation(true, objDoc);
                                CreateJournal(true, objDoc);
                                CreateCapitalization(objDoc);
                                break;
                            case BoObjectTypes.oPurchaseCreditNotes:
                                CreateRevaluation(false, objDoc);
                                CreateJournal(false, objDoc);
                                CreateCreditCapitalization(objDoc);
                                break;
                        }
                    }
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objDoc);
                GC.Collect();
            }
        }
        #endregion

        #region Métodos de negocio
        private static void CreateCapitalization(Documents oDoc)
         {
            AssetDocumentService oAssetService = (AssetDocumentService)MainObject.Instance.B1Company.GetCompanyService().GetBusinessService(ServiceTypes.AssetCapitalizationService);
            AssetDocumentParams faDocumentParams = (AssetDocumentParams)oAssetService.GetDataInterface(AssetDocumentServiceDataInterfaces.adsAssetDocumentParams);
            AssetDocument oAssetDocument = (SAPbobsCOM.AssetDocument)oAssetService.GetDataInterface(SAPbobsCOM.AssetDocumentServiceDataInterfaces.adsAssetDocument);
            SalesTaxCodes otaxCode = (SalesTaxCodes)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oSalesTaxCodes);
            SalesTaxAuthorities Ta = (SalesTaxAuthorities)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oSalesTaxAuthorities);
            SAPbobsCOM.Items oItem = (SAPbobsCOM.Items)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oItems);

            try
            {
                var thirdToUse = GetValueThird(oDoc.CardCode);
                oAssetDocument.AssetValueDate = oDoc.DocDate;
                oAssetDocument.Remarks = "Depreciación por IVA Costo Doc: " + oDoc.DocNum;
                oAssetDocument.DepreciationArea = "*";

                if (oDoc.DocType == BoDocumentTypes.dDocument_Items)
                {
                    for (int i = 0; i < oDoc.Lines.Count; i++)
                    {
                        oDoc.Lines.SetCurrentLine(i);
                        oItem.GetByKey(oDoc.Lines.ItemCode);
                        otaxCode.GetByKey(oDoc.Lines.TaxCode);
                        if (otaxCode.UserFields.Fields.Item("U_HCO_IvaCosto").Value.ToString().Equals("Y") && oItem.InventoryItem == BoYesNoEnum.tNO && oItem.ItemType == ItemTypeEnum.itFixedAssets)
                        {
                            Ta.GetByKey(otaxCode.Lines.STACode, otaxCode.Lines.STAType);
                            if (oItem.VirtualAssetItem == BoYesNoEnum.tYES)
                            {
                                for (int j = 0; j < oDoc.Lines.GeneratedAssets.Count; j++)
                                {
                                    oDoc.Lines.GeneratedAssets.SetCurrentLine(j);
                                    AssetDocumentLine oLine = oAssetDocument.AssetDocumentLineCollection.Add();
                                    oLine.AssetNumber = oDoc.Lines.GeneratedAssets.AssetCode;
                                    oLine.TotalLC = oDoc.Lines.TaxTotal / oDoc.Lines.Quantity;
                                }
                            }
                            else
                            {
                                AssetDocumentLine oLine = oAssetDocument.AssetDocumentLineCollection.Add();
                                oLine.AssetNumber = oDoc.Lines.ItemCode;
                                oLine.TotalLC = oDoc.Lines.TaxTotal / oDoc.Lines.Quantity;
                            }
                        }
                    }
                }
                var response = oAssetService.Add(oAssetDocument);
                ActualizarInfoCapitalizacion(response.Code.ToString(), thirdToUse);
            }
            catch (COMException comEx)
            {
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", comEx.ErrorCode, "::", comEx.Message, "::", comEx.StackTrace })));
                _Logger.Error("", exception);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oAssetService);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(faDocumentParams);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oAssetDocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(otaxCode);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Ta);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
                GC.Collect();
            }
        }

        public static void CreateIVAJournal(SAPbouiCOM.BusinessObjectInfo pVal)
        {
            var oXml = new XmlDocument();
            oXml.LoadXml(pVal.ObjectKey);

            var form = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            var strSQL = string.Format(Queries.Instance.Queries().Get("GetJournalPaymentNumber"), form.DataSources.DBDataSources.Item(0).TableName, oXml.InnerText);
            var objRS = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            objRS.DoQuery(strSQL);
            objRS.MoveFirst();

            if (objRS.RecordCount > 0)
            {
                var objDoc = (SAPbobsCOM.Documents)MainObject.Instance.B1Company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), pVal.Type));
                    objDoc.GetByKey(int.Parse(oXml.InnerText));

                var thirdToUse = GetValueThird(objDoc.CardCode);
                var queryJournal = string.Format(Queries.Instance.Queries().Get("GetValueToJournal"), objRS.Fields.Item("TransId").Value);
                var journal = (SAPbobsCOM.JournalEntries)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oJournalEntries);
                var objRSJournal = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objRSJournal.DoQuery(queryJournal);

                if (objRSJournal.RecordCount > 0)
                {
                    while (!objRSJournal.EoF)
                    {
                        journal.Lines.AccountCode = objRSJournal.Fields.Item("SalesTax").Value.ToString();
                        journal.Lines.Debit = double.Parse(objRSJournal.Fields.Item("Credit").Value.ToString());
                        journal.Lines.UserFields.Fields.Item("U_HCO_RELPAR").Value = thirdToUse;
                        journal.Lines.Add();

                        journal.Lines.AccountCode = objRSJournal.Fields.Item("U_HCO_CtaIva").Value.ToString();
                        journal.Lines.Credit = double.Parse(objRSJournal.Fields.Item("Credit").Value.ToString());
                        journal.Lines.UserFields.Fields.Item("U_HCO_RELPAR").Value = thirdToUse;
                        journal.Lines.Add();

                        objRSJournal.MoveNext();
                    }

                    journal.Add();
                }
            }
        }

        private static void CreateCreditCapitalization(Documents oDoc)
        {
            AssetDocumentService oAssetService = (AssetDocumentService)MainObject.Instance.B1Company.GetCompanyService().GetBusinessService(ServiceTypes.AssetCapitalizationCreditMemoService);
            AssetDocumentParams faDocumentParams = (AssetDocumentParams)oAssetService.GetDataInterface(AssetDocumentServiceDataInterfaces.adsAssetDocumentParams);
            AssetDocument oAssetDocument = (SAPbobsCOM.AssetDocument)oAssetService.GetDataInterface(SAPbobsCOM.AssetDocumentServiceDataInterfaces.adsAssetDocument);
            SalesTaxCodes otaxCode = (SalesTaxCodes)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oSalesTaxCodes);
            SalesTaxAuthorities Ta = (SalesTaxAuthorities)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oSalesTaxAuthorities);
            SAPbobsCOM.Items oItem = (SAPbobsCOM.Items)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oItems);

            try
            {

                oAssetDocument.AssetValueDate = oDoc.DocDate;
                oAssetDocument.Remarks = "Depreciación crédito por IVA Costo Doc: " + oDoc.DocNum;
                oAssetDocument.DepreciationArea = "*";

                if (oDoc.DocType == BoDocumentTypes.dDocument_Items)
                {
                    for (int i = 0; i < oDoc.Lines.Count; i++)
                    {
                        oDoc.Lines.SetCurrentLine(i);
                        oItem.GetByKey(oDoc.Lines.ItemCode);
                        otaxCode.GetByKey(oDoc.Lines.TaxCode);
                        if (otaxCode.UserFields.Fields.Item("U_HCO_IvaCosto").Value.ToString().Equals("Y") && oItem.InventoryItem == BoYesNoEnum.tNO && oItem.ItemType == ItemTypeEnum.itFixedAssets)
                        {
                            Ta.GetByKey(otaxCode.Lines.STACode, otaxCode.Lines.STAType);
                            if (oItem.VirtualAssetItem == BoYesNoEnum.tYES)
                            {
                                for (int j = 0; j < oDoc.Lines.GeneratedAssets.Count; j++)
                                {
                                    oDoc.Lines.GeneratedAssets.SetCurrentLine(j);
                                    AssetDocumentLine oLine = oAssetDocument.AssetDocumentLineCollection.Add();
                                    oLine.AssetNumber = oDoc.Lines.GeneratedAssets.AssetCode;
                                    oLine.TotalLC = oDoc.Lines.TaxTotal / oDoc.Lines.Quantity;
                                }
                            }
                            else
                            {
                                AssetDocumentLine oLine = oAssetDocument.AssetDocumentLineCollection.Add();
                                oLine.AssetNumber = oDoc.Lines.ItemCode;
                                oLine.TotalLC = oDoc.Lines.TaxTotal / oDoc.Lines.Quantity;
                            }
                        }
                    }
                }
                oAssetService.Add(oAssetDocument);
            }
            catch (COMException comEx)
            {
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", comEx.ErrorCode, "::", comEx.Message, "::", comEx.StackTrace })));
                _Logger.Error("", exception);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oAssetService);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(faDocumentParams);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oAssetDocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(otaxCode);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Ta);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
                GC.Collect();
            }
        }

        public static void CreateJournal(bool Debit, Documents oDoc)
        {
            JournalEntries oJE = (JournalEntries)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oJournalEntries);
            SalesTaxCodes otaxCode = (SalesTaxCodes)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oSalesTaxCodes);
            SalesTaxAuthorities Ta = (SalesTaxAuthorities)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oSalesTaxAuthorities);
            SAPbobsCOM.Items oItem = (SAPbobsCOM.Items)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oItems);
            bool count = false;

            try
            {
                var thirdToUse = GetValueThird(oDoc.CardCode);
                oJE.ReferenceDate = oDoc.DocDate;
                oJE.Memo = "Contabilización IVA Costo Doc: " + +oDoc.DocNum;

                if (oDoc.DocType == BoDocumentTypes.dDocument_Items)
                {
                    for (int i = 0; i < oDoc.Lines.Count; i++)
                    {
                        oDoc.Lines.SetCurrentLine(i);
                        oItem.GetByKey(oDoc.Lines.ItemCode);
                        otaxCode.GetByKey(oDoc.Lines.TaxCode);
                        if (otaxCode.UserFields.Fields.Item("U_HCO_IvaCosto").Value.ToString().Equals("Y") && oItem.InventoryItem == BoYesNoEnum.tNO)
                        {
                            count = true;
                            Ta.GetByKey(otaxCode.Lines.STACode, otaxCode.Lines.STAType);
                            oJE.Lines.AccountCode = Ta.AOrPTaxAccount;
                            if (Debit) oJE.Lines.Credit = oDoc.Lines.TaxTotal;
                            else oJE.Lines.Debit = oDoc.Lines.TaxTotal;
                            oJE.Lines.CostingCode = oDoc.Lines.CostingCode;
                                oJE.Lines.CostingCode2 = oDoc.Lines.CostingCode2;
                            oJE.Lines.CostingCode3 = oDoc.Lines.CostingCode3;
                            oJE.Lines.CostingCode4 = oDoc.Lines.CostingCode4;
                            oJE.Lines.CostingCode5 = oDoc.Lines.CostingCode5;
                            oJE.Lines.UserFields.Fields.Item("U_HCO_RELPAR").Value = thirdToUse;
                            oJE.Lines.Add();
                            oJE.Lines.AccountCode = oDoc.Lines.AccountCode;
                            if (Debit) oJE.Lines.Debit = oDoc.Lines.TaxTotal;
                            else oJE.Lines.Credit = oDoc.Lines.TaxTotal;
                            oJE.Lines.CostingCode = oDoc.Lines.CostingCode;
                            oJE.Lines.CostingCode2 = oDoc.Lines.CostingCode2;
                            oJE.Lines.CostingCode3 = oDoc.Lines.CostingCode3;
                            oJE.Lines.CostingCode4 = oDoc.Lines.CostingCode4;
                            oJE.Lines.CostingCode5 = oDoc.Lines.CostingCode5;
                            oJE.Lines.UserFields.Fields.Item("U_HCO_RELPAR").Value = thirdToUse;
                            oJE.Lines.Add();
                        }
                    }
                }
                else
                {
                    count = true;
                    for (int i = 0; i < oDoc.Lines.Count; i++)
                    {
                        oDoc.Lines.SetCurrentLine(i);
                        otaxCode.GetByKey(oDoc.Lines.TaxCode);
                        if (otaxCode.UserFields.Fields.Item("U_HCO_IvaCosto").Value.ToString().Equals("Y"))
                        {

                            Ta.GetByKey(otaxCode.Lines.STACode, otaxCode.Lines.STAType);
                            oJE.Lines.AccountCode = Ta.AOrPTaxAccount;
                            if (Debit) oJE.Lines.Credit = oDoc.Lines.TaxTotal;
                            else oJE.Lines.Debit = oDoc.Lines.TaxTotal;
                            oJE.Lines.CostingCode = oDoc.Lines.CostingCode;
                            oJE.Lines.CostingCode2 = oDoc.Lines.CostingCode2;
                            oJE.Lines.CostingCode3 = oDoc.Lines.CostingCode3;
                            oJE.Lines.CostingCode4 = oDoc.Lines.CostingCode4;
                            oJE.Lines.CostingCode5 = oDoc.Lines.CostingCode5;
                            oJE.Lines.UserFields.Fields.Item("U_HCO_RELPAR").Value = thirdToUse;
                            oJE.Lines.Add();
                            oJE.Lines.AccountCode = oDoc.Lines.AccountCode;
                            if (Debit) oJE.Lines.Debit = oDoc.Lines.TaxTotal;
                            else oJE.Lines.Credit = oDoc.Lines.TaxTotal;
                            oJE.Lines.CostingCode = oDoc.Lines.CostingCode;
                            oJE.Lines.CostingCode2 = oDoc.Lines.CostingCode2;
                            oJE.Lines.CostingCode3 = oDoc.Lines.CostingCode3;
                            oJE.Lines.CostingCode4 = oDoc.Lines.CostingCode4;
                            oJE.Lines.CostingCode5 = oDoc.Lines.CostingCode5;
                            oJE.Lines.UserFields.Fields.Item("U_HCO_RELPAR").Value = thirdToUse;
                            oJE.Lines.Add();

                        }
                    }
                }

                if (count)
                {
                    if (oJE.Add() != 0)
                    {
                        MainObject.Instance.B1Application.SetStatusBarMessage(MainObject.Instance.B1Company.GetLastErrorDescription(), BoMessageTime.bmt_Short, true);
                    }
                }
            }
            catch (COMException comEx)
            {
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", comEx.ErrorCode, "::", comEx.Message, "::", comEx.StackTrace })));
                _Logger.Error("", exception);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oJE);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(otaxCode);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Ta);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
                GC.Collect();
            }

        }

        static public void CheckCapitalizationGenerated(string docEntry, Documents oDoc)
        {
            var thirdToUse = GetValueThird(oDoc.CardCode);
            var queryJournal = string.Format(Queries.Instance.Queries().Get("GetCapitalizationValue"), docEntry);
            var objRSJournal = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objRSJournal.DoQuery(queryJournal);

            if( objRSJournal.RecordCount > 0 )
            {
                while(!objRSJournal.EoF)
                {
                    ActualizarInfoCapitalizacion(objRSJournal.Fields.Item("DocEntry").Value.ToString(), thirdToUse);
                    objRSJournal.MoveNext();
                }
            }
        }

        static public void ActualizarInfoCapitalizacion(string docEntry, string thirdToUse)
        {
            UpdateJournalCapitalization(docEntry, thirdToUse);
        }

        static private void UpdateJournalCapitalization(string docEntry, string thirdToUse)
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
                    for( int j=0; j < journal.Lines.Count; j++)
                    {
                        journal.Lines.SetCurrentLine(j);
                        journal.Lines.UserFields.Fields.Item("U_HCO_RELPAR").Value = thirdToUse;
                    }
                    journal.Update();
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

        static public string GetValueThird(string cardCode)
        {
            try
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
            catch
            {
                return String.Empty;
            }
        }

        public static void CreateRevaluation(bool Debit, Documents oDoc)
        {
            if (oDoc.DocType == BoDocumentTypes.dDocument_Service) return;
            if (oDoc.Lines.BaseType == 20 && oDoc.DocObjectCode == BoObjectTypes.oPurchaseInvoices) return;
            if (oDoc.Lines.BaseType == 18 && oDoc.DocObjectCode == BoObjectTypes.oPurchaseCreditNotes)
            {
                Documents oInv = (Documents)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oPurchaseInvoices);
                oInv.GetByKey(oDoc.Lines.BaseEntry);
                if (oInv.Lines.BaseType == 20) return;

            }

            var thirdToUse = GetValueThird(oDoc.CardCode);
            MaterialRevaluation oRev = (MaterialRevaluation)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oMaterialRevaluation);
            SalesTaxCodes otaxCode = (SalesTaxCodes)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oSalesTaxCodes);
            SalesTaxAuthorities Ta = (SalesTaxAuthorities)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oSalesTaxAuthorities);
            SAPbobsCOM.Items oItem = (SAPbobsCOM.Items)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oItems);

            try
            {
                oRev.RevalType = "M";
                oRev.DocDate = oDoc.DocDate;
                oRev.Comments = "Revalorización por IVA Costo Doc: " + oDoc.DocNum;
                int count = 0;

                for (int i = 0; i < oDoc.Lines.Count; i++)
                {
                    oDoc.Lines.SetCurrentLine(i);
                    oItem.GetByKey(oDoc.Lines.ItemCode);
                    otaxCode.GetByKey(oDoc.Lines.TaxCode);
                    if (otaxCode.UserFields.Fields.Item("U_HCO_IvaCosto").Value.ToString().Equals("Y") && oItem.InventoryItem == BoYesNoEnum.tYES)
                    {
                        Ta.GetByKey(otaxCode.Lines.STACode, otaxCode.Lines.STAType);
                        oRev.Lines.SetCurrentLine(count);
                        oRev.Lines.ItemCode = oDoc.Lines.ItemCode;
                        oRev.Lines.WarehouseCode = oDoc.Lines.WarehouseCode;
                        oRev.Lines.Quantity = oDoc.Lines.Quantity;
                        oRev.Lines.DebitCredit = Debit ? oDoc.Lines.TaxTotal : (oDoc.Lines.TaxTotal * -1);
                        if (Debit) oRev.Lines.RevaluationIncrementAccount = Ta.AOrPTaxAccount;
                        else oRev.Lines.RevaluationDecrementAccount = Ta.AOrPTaxAccount;
                        oRev.Lines.Add();
                        count++;
                    }
                }
                if (count > 0)
                {
                    if (oRev.Add() != 0)
                    {
                        var msn = MainObject.Instance.B1Company.GetLastErrorDescription();
                    }
                    else
                    {
                        var lastRevalDocEntry = MainObject.Instance.B1Company.GetNewObjectKey();
                        var query = string.Format(Queries.Instance.Queries().Get("GetRevaluationJournal"), lastRevalDocEntry);
                        var record = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                            record.DoQuery(query);
                        if( record.RecordCount > 0 )
                        {
                            var journal =(JournalEntries) MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oJournalEntries);
                            if( journal.GetByKey(int.Parse(record.Fields.Item("TransId").Value.ToString())) )
                            {
                                for(int i=0; i < journal.Lines.Count; i++)
                                {
                                    journal.Lines.SetCurrentLine(i);
                                    journal.Lines.UserFields.Fields.Item("U_HCO_RELPAR").Value = thirdToUse;
                                }

                                journal.Update();
                            }
                        }
                    }
                }
            }
            catch (COMException comEx)
            {
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", comEx.ErrorCode, "::", comEx.Message, "::", comEx.StackTrace })));
                _Logger.Error("", exception);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRev);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(otaxCode);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Ta);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
                GC.Collect();
            }
        }

        private static bool HasIVACosto(SAPbobsCOM.Documents oDoc)
        {
            Recordset oRS = (SAPbobsCOM.Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            string DocType = string.Empty;

            switch (oDoc.DocObjectCode)
            {
                case BoObjectTypes.oPurchaseInvoices:
                    DocType = "PCH1";
                    break;
                case BoObjectTypes.oPurchaseCreditNotes:
                    DocType = "RPC1";
                    break;
                case BoObjectTypes.oPurchaseDeliveryNotes:
                    DocType = "PDN1";
                    break;
                case BoObjectTypes.oPurchaseReturns:
                    DocType = "RPD1";
                    break;
            }

            string query = string.Format(Queries.Instance.Queries().Get("HasIVACosto"), DocType, oDoc.DocEntry);

            try
            {
                oRS.DoQuery(query);
                if (oRS.RecordCount > 0)
                {
                    oRS.MoveFirst();
                    return Int32.Parse(oRS.Fields.Item(0).Value.ToString()) > 0;
                }
            }
            catch (COMException comEx)
            {
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", comEx.ErrorCode, "::", comEx.Message, "::", comEx.StackTrace })));
                _Logger.Error("", exception);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                GC.Collect();
            }

            return false;
        }
        #endregion

    }
}
