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
using System.Reflection;

namespace T1.B1.WithholdingTax
{
    public class WithholdingTax
    {
        private static WithholdingTax objWithHoldingTax;
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static Recordset objRS;
        private static DBDataSource objDBDataSource;
        private static string strSQL;
        private static XmlDocument oXml;
        private static CompanyService objCompany = null;
        private static GeneralService objGeneralService = null;
        private static GeneralDataParams objGenParams = null;
        private static GeneralData objGeneralData = null;
        private static GeneralDataCollection objGenDataColl = null;

        private WithholdingTax()
        {
            if (objWithHoldingTax == null) objWithHoldingTax = new WithholdingTax();
        }

        #region Form
        public static string GetXmlUDO(Settings.EWithHoldingTax type)
        {
            XmlDocument oXML = new XmlDocument();
            switch (type)
            {
                case Settings.EWithHoldingTax.MUNICIPALITY:
                    oXML.Load(AppDomain.CurrentDomain.BaseDirectory + "\\Forms\\GruposMunicipios.srf");
                    break;
                case Settings.EWithHoldingTax.WITHHOLDING_OPERATION:
                    oXML.Load(AppDomain.CurrentDomain.BaseDirectory + "\\Forms\\RegistroRetenciones.srf");
                    break;
                case Settings.EWithHoldingTax.MISSING_OPERATIONS:
                    oXML.Load(AppDomain.CurrentDomain.BaseDirectory + "\\Forms\\OperacionesFaltantesRetencion.srf");
                    break;
                default:
                    return string.Empty;
            }
            return oXML.InnerXml;
        }
        public static string LoadWithHoldingForm(Settings.EWithHoldingTax type)
        {
            string formId = string.Empty;
            try
            {
                FormCreationParams objParams = (FormCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
                objParams.XmlData = GetXmlUDO(type);
                objParams.FormType = GetTypeUDO(type);
                objParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);

                Form objForm = MainObject.Instance.B1Application.Forms.AddEx(objParams);
                objForm.VisibleEx = true;
                formId = objForm.UniqueID;

            }
            catch (Exception er)
            {
                _Logger.Error("(LoadWithHoldingForm)", er);
            }
            finally
            {
                GC.Collect();
            }

            return formId;
        }
        public static string GetTypeUDO(Settings.EWithHoldingTax type)
        {
            switch (type)
            {
                case Settings.EWithHoldingTax.MUNICIPALITY:
                    return "HCO_FWT0100";
                case Settings.EWithHoldingTax.WITHHOLDING_OPERATION:
                    return "HCO_FWT1100";
                case Settings.EWithHoldingTax.MISSING_OPERATIONS:
                    return "HCO_FWT1200";
                default:
                    return string.Empty;
            }
        }
        internal static void InitMissingOperationsForm(string FormId)
        {
            SAPbouiCOM.DataTable objDT;
            Matrix objMatrix;
            Form objForm = MainObject.Instance.B1Application.Forms.Item(FormId);
            try
            {
                objForm.Freeze(true);
                objDT = objForm.DataSources.DataTables.Item("DT_TRA");
                objDT.ExecuteQuery(Queries.Instance.Queries().Get("GetMissingOperations"));
                objMatrix = (Matrix)objForm.Items.Item("grTRA").Specific;
                objMatrix.Columns.Item("#").DataBind.Bind("DT_TRA", "LineId");
                objMatrix.Columns.Item("Col_0").DataBind.Bind("DT_TRA", "Seleccionar");
                objMatrix.Columns.Item("Col_1").DataBind.Bind("DT_TRA", "Documento");
                objMatrix.Columns.Item("Col_2").DataBind.Bind("DT_TRA", "Numero");
                objMatrix.Columns.Item("Col_3").DataBind.Bind("DT_TRA", "CardCode");
                objMatrix.Columns.Item("Col_4").DataBind.Bind("DT_TRA", "DocTotal");
                objMatrix.Columns.Item("Col_5").DataBind.Bind("DT_TRA", "DocType");
                objMatrix.Columns.Item("Col_5").Visible = false;
                objMatrix.Columns.Item("Col_6").DataBind.Bind("DT_TRA", "DocEntry");
                objMatrix.Columns.Item("Col_6").Visible = false;
                objMatrix.LoadFromDataSource();
                objMatrix.AutoResizeColumns();
                objForm.Freeze(false);

            }
            catch (Exception er)
            {
                _Logger.Error("(LoadWithHoldingForm)", er);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objForm);
                GC.Collect();
            }

        }
        internal static void InitBusinessPartnerForm(string FormId)
        {
            Form objForm = MainObject.Instance.B1Application.Forms.Item(FormId);
            try
            {
                objForm.Freeze(true);
                LinkedButton oLink = (LinkedButton)objForm.Items.Item("Item_13").Specific;
                oLink.LinkedObject = (BoLinkedObject)(Int32.Parse(objForm.DataSources.DBDataSources.Item(0).GetValue("U_DocType", 0)));
                oLink.Item.Visible = true;
                oLink = (LinkedButton)objForm.Items.Item("Item_11").Specific;
                oLink.Item.Visible = true;
                objForm.Freeze(false);
            }
            catch (Exception er)
            {
                _Logger.Error("(InitBusinessPartnerForm)", er);
                objForm.Freeze(false);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objForm);
                GC.Collect();
            }
        }
        internal static void InitWithHoldingOperForm(string FormId)
        {
            Form objForm = MainObject.Instance.B1Application.Forms.Item(FormId);
            try
            {
                LinkedButton oLink = (LinkedButton)objForm.Items.Item("Item_11").Specific;
                oLink.LinkedObjectType = "HCO_FRP1100";
                oLink.LinkedFormXmlPath = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Forms\\HCO_Terceros_Relacionados.srf";
                oLink.Item.Visible = false;
            }
            catch (Exception er)
            {
                _Logger.Error("(InitWithHoldingOperForm)", er);
                objForm.Freeze(false);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objForm);
                GC.Collect();
            }
        }

        #endregion


        static public bool GetSelectedBPInformation(Form form, bool reset)
        {
            //bool blReadWTConfig = false;
            bool b1AddOnCalc = false;
            string strLastCardCode = string.Empty;
            string strPickedCardCode = string.Empty;
            int basetype = 0, baseEntry = 0;
            Form objForm;

            try
            {
                objForm = form;
                bool isDisabled = CacheManager.CacheManager.Instance.getFromCache("Disable_" + objForm.UniqueID) == null ? false : true;
                if (!isDisabled)
                {
                    strLastCardCode = CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + objForm.UniqueID) == null ? "" : CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + objForm.UniqueID);
                    //blAddOnCalc = CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + FormId) == null ? false : true;
                    strPickedCardCode = objForm.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0).Trim();

                    if ((strLastCardCode.Trim() != strPickedCardCode.Trim() && !strPickedCardCode.Trim().Equals("")) || reset)
                    {
                        if (!WithholdingTax.IsBasedNC(objForm, ref basetype, ref baseEntry))
                            WithholdingTax.GetWTfromBP(strPickedCardCode, objForm);
                        else
                            WithholdingTax.GetWTfromDOC(objForm, basetype, baseEntry);

                        CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + objForm.UniqueID, strPickedCardCode, CacheManager.CacheManager.objCachePriority.Default);
                        b1AddOnCalc = true;
                    }
                    else if (CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + objForm.UniqueID) != null ? true : false)
                    {
                        UpdateBases(objForm);
                        if (!IsBasedNC(objForm.UniqueID))
                            b1AddOnCalc = isUpdateble(CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTFormInfoCachePrefix + objForm.UniqueID));
                        else
                            b1AddOnCalc = true;
                    }
                }
            }
            catch (COMException comEx)
            {
                _Logger.Error("", comEx);
                MainObject.Instance.B1Application.SetStatusBarMessage("HCO: " + comEx.InnerException.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("HCO: " + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }

            return b1AddOnCalc;
        }
        public static bool IsBasedNC(Form form, ref int baseType, ref int docEntry)
        {
            Form objForm = form;
            bool result = false;
            EditText edt = (EditText)((Matrix)objForm.Items.Item("38").Specific).GetCellSpecific("43", 1);
            try
            {
                baseType = int.Parse(edt.Value.ToString());

                if (baseType == 13 || baseType == 18)
                {
                    edt = (EditText)((Matrix)objForm.Items.Item("38").Specific).GetCellSpecific("45", 1);
                    docEntry = int.Parse(edt.Value.ToString());
                    result = true;
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
                GC.Collect();
            }

            return result;
        }
        public static bool IsBasedNC(string formId)
        {
            Form objForm = MainObject.Instance.B1Application.Forms.Item(formId);
            int baseType = 0;

            try
            {
                baseType = int.Parse(objForm.DataSources.DBDataSources.Item(10).GetValue("BaseType", 0));
                return (baseType == 13 || baseType == 18);
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
                GC.Collect();
            }

            return false;
        }

        public static void GetWTfromDOC(Form form, int baseType, int docEntry)
        {
            Form objForm;
            List<WithholdingTaxDetail> WTDocInfo;
            Documents objDOC = null;
            WithholdingTaxCodes objWTInfo = null;
            WTDocInfo = new List<WithholdingTaxDetail>();

            try
            {
                objForm = form;
                objDOC = (Documents)MainObject.Instance.B1Company.GetBusinessObject(baseType == 13 ? BoObjectTypes.oInvoices : BoObjectTypes.oPurchaseInvoices);
                objWTInfo = (WithholdingTaxCodes)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oWithholdingTaxCodes);

                if (objDOC.GetByKey(docEntry))
                {

                    for (int i = 0; i < objDOC.WithholdingTaxData.Count; i++)
                    {
                        objDOC.WithholdingTaxData.SetCurrentLine(i);
                        if (objWTInfo.GetByKey(objDOC.WithholdingTaxData.WTCode))
                        {
                            if (objWTInfo.Inactive == BoYesNoEnum.tNO)
                            {
                                WithholdingTaxDetail objDet = new WithholdingTaxDetail();
                                objDet.WTCode = objDOC.WithholdingTaxData.WTCode;
                                objDet.Rate = objWTInfo.BaseAmount;
                                objDet.MMCode = objWTInfo.UserFields.Fields.Item("U_HCO_MMCode").Value.ToString();
                                objDet.Area = objWTInfo.UserFields.Fields.Item("U_HCO_Area").Value.ToString();
                                objDet.MinBase = double.Parse(objWTInfo.UserFields.Fields.Item("U_HCO_MinBase").Value.ToString());
                                objDet.WTType = int.Parse(objWTInfo.UserFields.Fields.Item("U_HCO_WTType").Value.ToString());
                                WTDocInfo.Add(objDet);
                            }
                        }
                    }
                }

                CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTFormInfoCachePrefix + objForm.UniqueID, WTDocInfo, CacheManager.CacheManager.objCachePriority.Default);
                CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + objForm.UniqueID, true, CacheManager.CacheManager.objCachePriority.Default);

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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objDOC);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objWTInfo);
                GC.Collect();
            }
        }
        public static void GetWTfromBP(string cardCode, Form form)
        {
            Form objForm;
            List<WithholdingTaxDetail> WTDocInfo;
            BusinessPartners objBP = null;
            WithholdingTaxCodes objWTInfo = null;
            string munCode = "";
            WTDocInfo = new List<WithholdingTaxDetail>();
            try
            {
                objForm = form;
                objBP = (BusinessPartners)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                objWTInfo = (WithholdingTaxCodes)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oWithholdingTaxCodes);

                if (objBP.GetByKey(cardCode))
                {
                    for (int i = 0; i < objBP.Addresses.Count; i++)
                    {
                        objBP.Addresses.SetCurrentLine(i);
                        if (objBP.CardType == BoCardTypes.cCustomer &&
                            objBP.Addresses.AddressName.Equals(objForm.DataSources.DBDataSources.Item(0).GetValue("ShipToCode", 0)) &&
                            objBP.Addresses.AddressType == BoAddressType.bo_ShipTo)
                            munCode = objBP.Addresses.UserFields.Fields.Item("U_HCO_MUNI").Value.ToString();
                        else if (objBP.CardType == BoCardTypes.cSupplier &&
                            objBP.Addresses.AddressName.Equals(objForm.DataSources.DBDataSources.Item(0).GetValue("PayToCode", 0)) &&
                            objBP.Addresses.AddressType == BoAddressType.bo_BillTo)
                            munCode = objBP.Addresses.UserFields.Fields.Item("U_HCO_MUNI").Value.ToString();
                    }
                    for (int i = 0; i < objBP.BPWithholdingTax.Count; i++)
                    {
                        objBP.BPWithholdingTax.SetCurrentLine(i);
                        if (objWTInfo.GetByKey(objBP.BPWithholdingTax.WTCode))
                        {
                            if (objWTInfo.Inactive == BoYesNoEnum.tNO)
                            {
                                WithholdingTaxDetail objDet = new WithholdingTaxDetail();
                                objDet.WTCode = objBP.BPWithholdingTax.WTCode;
                                objDet.Rate = objWTInfo.BaseAmount;
                                objDet.MMCode = objWTInfo.UserFields.Fields.Item("U_HCO_MMCode").Value.ToString();
                                objDet.Area = objWTInfo.UserFields.Fields.Item("U_HCO_Area").Value.ToString();
                                objDet.MinBase = double.Parse(objWTInfo.UserFields.Fields.Item("U_HCO_MinBase").Value.ToString());
                                objDet.MunGroup = objWTInfo.UserFields.Fields.Item("U_HCO_MunGroup").Value.ToString();
                                objDet.WTType = int.Parse(objWTInfo.UserFields.Fields.Item("U_HCO_WTType").Value.ToString());
                                objDet.Municipios = GetWTMuniInfo(objDet.MunGroup);

                                if (objDet.MunGroup.Equals(string.Empty)) WTDocInfo.Add(objDet);
                                else
                                {
                                    foreach (WithholdingTaxConfigMun mun in objDet.Municipios)
                                    {
                                        if (mun.MunCode.Equals(munCode)) WTDocInfo.Add(objDet);
                                    }
                                }
                            }
                        }
                    }
                }

                CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTFormInfoCachePrefix + objForm.UniqueID, WTDocInfo, CacheManager.CacheManager.objCachePriority.Default);
                CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + objForm.UniqueID, true, CacheManager.CacheManager.objCachePriority.Default);

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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objBP);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objWTInfo);
                GC.Collect();
            }
        }

        public static void UpdateBases(Form form)
        {
            Form objForm;
            List<WithholdingTaxDetail> WTDocInfo;
            double NetBase = 0;
            double VatBase = 0;

            objForm = form;
            try
            {
                if (CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + objForm.UniqueID) != null ? true : false)
                {
                    WTDocInfo = CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTFormInfoCachePrefix + objForm.UniqueID);
                    GetMatrixBaseAmount(((Matrix)objForm.Items.Item("38").Specific), ref NetBase, ref VatBase, ((ComboBox)objForm.Items.Item("63").Specific).Value);

                    foreach (WithholdingTaxDetail oDet in WTDocInfo)
                    {
                        oDet.NetBase = NetBase;
                        oDet.VatBase = VatBase;
                    }

                    CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTFormInfoCachePrefix + objForm.UniqueID, WTDocInfo, CacheManager.CacheManager.objCachePriority.Default);
                }
            }
            catch (COMException comEx)
            {
                _Logger.Error("", comEx);
                MainObject.Instance.B1Application.SetStatusBarMessage("HCO: " + comEx.InnerException.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("HCO: " + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }
        public static bool isUpdateble(List<WithholdingTaxDetail> WTDocInfo)
        {
            bool updateble = false;
            foreach (WithholdingTaxDetail oDet in WTDocInfo)
            {
                updateble = (oDet.isMinBaseValid && !oDet.assigned) || (!oDet.isMinBaseValid && oDet.assigned);
                if (updateble) break;
            }
            return updateble;
        }

        public static void SetTypeWT(string FormUID_WT, string FormUID)
        {
            if (!IsBasedNC(FormUID)) SetBPWT(FormUID_WT, FormUID);
            else SetDocWT(FormUID_WT, FormUID);
        }
        public static void SetBPWT(string FormUID_WT, string FormUID)
        {
            Form objWTForm = null;
            Form objForm = null;
            List<WithholdingTaxDetail> WTDocInfo;
            Matrix objMatrix = null;
            EditText objEdit = null;
            EditText objEditBase = null;
            int intNum = -1;
            double NetBase = 0, VatBase = 0;

            objForm = MainObject.Instance.B1Application.Forms.Item(FormUID);

            try
            {
                objWTForm = MainObject.Instance.B1Application.Forms.Item(FormUID_WT);
                WithholdingTax.GetBaseAmount(objForm.DataSources.DBDataSources.Item(10), ref NetBase, ref VatBase);

                if (CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTFormInfoCachePrefix + FormUID) != null)
                {
                    WTDocInfo = CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTFormInfoCachePrefix + FormUID);
                    intNum = 1;
                    objMatrix = (Matrix)objWTForm.Items.Item("6").Specific;
                    objMatrix.Clear();

                    foreach (WithholdingTaxDetail oDetail in WTDocInfo)
                    {
                        CacheManager.CacheManager.Instance.addToCache(string.Concat("Updating_", FormUID), true, CacheManager.CacheManager.objCachePriority.Default);
                        oDetail.NetBase = NetBase;
                        oDetail.VatBase = VatBase;

                        if (oDetail.isMinBaseValid && NetBase + VatBase != 0)
                        {
                            objMatrix.AddRow(1, -1);
                            objEdit = (EditText)objMatrix.GetCellSpecific("1", intNum);
                            objEditBase = (EditText)objMatrix.GetCellSpecific("U_HCO_BaseAmnt", intNum);
                            objEdit.Value = oDetail.WTCode;
                            objEditBase.Value = oDetail.WTType == 1 ? oDetail.VatBase.ToString(System.Globalization.CultureInfo.InvariantCulture) : oDetail.NetBase.ToString(System.Globalization.CultureInfo.InvariantCulture);
                            intNum++;
                            oDetail.assigned = true;
                        }
                        else oDetail.assigned = false;
                    }
                    
                    CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTFormInfoCachePrefix + objForm.UniqueID, WTDocInfo, CacheManager.CacheManager.objCachePriority.Default);
                }
                string strFormAutoActivate = CacheManager.CacheManager.Instance.getFromCache("WTAutoActivate") != null ? CacheManager.CacheManager.Instance.getFromCache("WTAutoActivate") : "";

                if (objWTForm.Mode != BoFormMode.fm_OK_MODE) objWTForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                CacheManager.CacheManager.Instance.removeFromCache("Updating_" + FormUID);
                if (strFormAutoActivate.Trim() == FormUID.Trim()) objWTForm.Close();

            }
            catch (COMException cOMException1)
            {
                _Logger.Error("", cOMException1);
            }
            catch (Exception exception2)
            {
                _Logger.Error("", exception2);
            }
        }
        public static void SetDocWT(string FormUID_WT, string FormUID)
        {
            Form objWTForm = null;
            Form objForm = null;
            List<WithholdingTaxDetail> WTDocInfo;
            Matrix objMatrix = null;
            EditText objEdit = null;
            EditText objEditBase = null;
            EditText objEditWT = null;
            double NetBase = 0, VatBase = 0;

            objForm = MainObject.Instance.B1Application.Forms.Item(FormUID);
            string strFormAutoActivate = CacheManager.CacheManager.Instance.getFromCache("WTAutoActivate") != null ? CacheManager.CacheManager.Instance.getFromCache("WTAutoActivate") : "";
            string decimalSeparator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;

            try
            {
                objWTForm = MainObject.Instance.B1Application.Forms.Item(FormUID_WT);
                GetMatrixBaseAmount(((Matrix)objForm.Items.Item("38").Specific), ref NetBase, ref VatBase, ((ComboBox)objForm.Items.Item("63").Specific).Value);

                if (CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTFormInfoCachePrefix + FormUID) != null)
                {
                    CacheManager.CacheManager.Instance.addToCache(string.Concat("Updating_", FormUID), true, CacheManager.CacheManager.objCachePriority.Default);
                    WTDocInfo = CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTFormInfoCachePrefix + FormUID);
                    objMatrix = (Matrix)objWTForm.Items.Item("6").Specific;

                    for (int i = 1; i <= objMatrix.RowCount; i++)
                    {
                        foreach (WithholdingTaxDetail oDetail in WTDocInfo)
                        {
                            oDetail.NetBase = NetBase;
                            oDetail.VatBase = VatBase;
                            objEdit = (EditText)objMatrix.GetCellSpecific("1", i);

                            if (objEdit.Value == oDetail.WTCode)
                            {
                                objEditBase = (EditText)objMatrix.GetCellSpecific("U_HCO_BaseAmnt", i);
                                objEditWT = (EditText)objMatrix.GetCellSpecific("14", i);
                                objEditBase.Value = oDetail.WTType == 1 ? oDetail.VatBase.ToString() : oDetail.NetBase.ToString();
                                objEditWT.Value = (oDetail.WTType == 1 ? (oDetail.VatBase * (oDetail.Rate / 100)).ToString() : (oDetail.NetBase * (oDetail.Rate / 100)).ToString()).Replace(decimalSeparator, MainObject.Instance.B1AdminInfo.DecimalSeparator);
                            }
                        }
                    }                   
                }

                if (objWTForm.Mode != BoFormMode.fm_OK_MODE) objWTForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                CacheManager.CacheManager.Instance.removeFromCache("Updating_" + FormUID);
                if (strFormAutoActivate.Trim() == FormUID.Trim()) objWTForm.Close();

            }
            catch (COMException cOMException1)
            {
                _Logger.Error("", cOMException1);
            }
            catch (Exception exception2)
            {
                _Logger.Error("", exception2);
            }
        }

        public static void GetBaseAmount(SAPbouiCOM.DBDataSource sap_table, ref double NetBase, ref double VatBase)
        {
            var XDoc = System.Xml.Linq.XDocument.Parse(sap_table.GetAsXML());
            var Rows = (XDoc.Element("dbDataSources").Element("rows").Elements("row")).ToList();

            if (!XDoc.ToString().Contains("WtLiable")) return;


            foreach (var Row in Rows)
            {
                string wtliable = (from h in Row.Descendants("cell")
                                   where h.Element("uid").Value == "WtLiable"
                                   select new
                                   {
                                       uid = h.Element("uid").Value,
                                       value = h.Element("value").Value
                                   }).First().value;

                if (wtliable.Equals("Y"))
                {
                    NetBase += double.Parse((from h in Row.Descendants("cell")
                                             where h.Element("uid").Value == "LineTotal"
                                             select new
                                             {
                                                 uid = h.Element("uid").Value,
                                                 value = h.Element("value").Value
                                             }).First().value, System.Globalization.CultureInfo.CurrentCulture);
                    VatBase += double.Parse((from h in Row.Descendants("cell")
                                             where h.Element("uid").Value == "VatSum"
                                             select new
                                             {
                                                 uid = h.Element("uid").Value,
                                                 value = h.Element("value").Value
                                             }).First().value, System.Globalization.CultureInfo.CurrentCulture);
                }

            }
        }
        public static void GetMatrixBaseAmount(SAPbouiCOM.Matrix sap_table, ref double NetBase, ref double VatBase, string docCurr)
        {
            var XDoc = System.Xml.Linq.XDocument.Parse(sap_table.SerializeAsXML(BoMatrixXmlSelect.mxs_All));
            var Rows = (XDoc.Element("Matrix").Element("Rows").Elements("Row")).ToList();
            string auxValue = string.Empty;

            string decimalSeparator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;

            foreach (var Row in Rows)
            {
                string wtliable = (from h in Row.Descendants("Column")
                                   where h.Element("ID").Value == "174"
                                   select new
                                   {
                                       uid = h.Element("ID").Value,
                                       value = h.Element("Value").Value
                                   }).First().value;

                if (wtliable.Equals("Y"))
                {
                    auxValue = (from h in Row.Descendants("Column")
                                where h.Element("ID").Value == "21"
                                select new
                                {
                                    uid = h.Element("ID").Value,
                                    value = h.Element("Value").Value
                                }).First().value;

                    auxValue = auxValue.Replace(docCurr, "").Trim();
                    auxValue = auxValue.Replace(MainObject.Instance.B1AdminInfo.ThousandsSeparator, "");
                    auxValue = auxValue.Replace(MainObject.Instance.B1AdminInfo.DecimalSeparator, decimalSeparator);

                    NetBase += double.Parse(auxValue, System.Globalization.CultureInfo.CurrentCulture);

                    auxValue = (from h in Row.Descendants("Column")
                                where h.Element("ID").Value == "82"
                                select new
                                {
                                    uid = h.Element("ID").Value,
                                    value = h.Element("Value").Value
                                }).First().value;

                    auxValue = auxValue.Replace(docCurr, "").Trim();
                    auxValue = auxValue.Replace(MainObject.Instance.B1AdminInfo.ThousandsSeparator, "");
                    auxValue = auxValue.Replace(MainObject.Instance.B1AdminInfo.DecimalSeparator, decimalSeparator);

                    VatBase += double.Parse(auxValue, System.Globalization.CultureInfo.CurrentCulture);
                }

            }
        }


        private static List<WithholdingTaxConfigMun> GetWTMuniInfo(string strCode)
        {
            List<WithholdingTaxConfigMun> objList = new List<WithholdingTaxConfigMun>();

            try
            {
                if (strCode.Trim().Length > 0)
                {
                    objCompany = MainObject.Instance.B1Company.GetCompanyService();
                    objGeneralService = objCompany.GetGeneralService(Settings._WithHoldingTax.WTMuniInfoUDO);
                    objGenParams = (GeneralDataParams)objGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                    objGenParams.SetProperty("Code", strCode);
                    objGeneralData = objGeneralService.GetByParams(objGenParams);
                    objGenDataColl = objGeneralData.Child(Settings._WithHoldingTax.WTMuniInfoChildUDO);
                    if (objGenDataColl.Count > 0)
                    {
                        foreach (GeneralData oDef in objGenDataColl)
                        {
                            string strMunCode = oDef.GetProperty("U_MunCode").ToString();
                            if (strMunCode != null && strMunCode.Trim().Length > 0)
                            {
                                WithholdingTaxConfigMun oMun = new WithholdingTaxConfigMun() { MunCode = strMunCode };
                                objList.Add(oMun);
                            }
                        }
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objCompany);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objGeneralService);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objGenParams);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objGeneralData);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objGenDataColl);
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
                GC.Collect();
            }
            return objList;
        }


        static public void SetFormState(Form form, FORM_MODE mode)
        {
            var type = form.TypeEx;
            var itemsEnc = new string[] { "Item_0", "23_U_E", "Item_2", "22_U_E", "20_U_E", "Item_9", "1_U_E", "Item_6", "25_U_E" };
            if (mode == FORM_MODE.SEARCH)
            {
                form.DataSources.DBDataSources.Item(0).SetValue("U_DocType", 0, "");
                foreach (var element in itemsEnc)
                    form.Items.Item(element).Enabled = true;
            }
        }
        static public string GetRelPartyCodeFromCardCode(string cardcode)
        {
            SAPbobsCOM.Recordset objRecordSet = null;
            string _relPartyCode = string.Empty;
            string strSQL = string.Format(Queries.Instance.Queries().Get("GetThirdRelated"), cardcode);
            try
            {
                objRecordSet = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if (objRecordSet != null)
                {
                    objRecordSet.DoQuery(strSQL);
                    if (objRecordSet != null && objRecordSet.RecordCount > 0) _relPartyCode = objRecordSet.Fields.Item(0).Value.ToString();
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            return _relPartyCode;
        }
        static public string GetCountyName(string code)
        {
            SAPbobsCOM.Recordset objRecordSet = null;
            string name = string.Empty;
            string strSQL = string.Format(Queries.Instance.Queries().Get("GetCountyName"), code);
            try
            {
                objRecordSet = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if (objRecordSet != null)
                {
                    objRecordSet.DoQuery(strSQL);
                    if (objRecordSet != null && objRecordSet.RecordCount > 0)
                    {
                        name = objRecordSet.Fields.Item(0).Value.ToString();

                    }
                }

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            return name;
        }

        public static void SetChooseFromListMunMatrix(ItemEvent pVal)
        {
            var oForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            var DbData = oForm.DataSources.DBDataSources.Item(1);
            var matrix = ((Matrix)oForm.Items.Item(pVal.ItemUID).Specific);
            var MunCode = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Code")[0].ToString();
            var MunName = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Name")[0].ToString();

            if (MunCode.Equals(string.Empty)) return;

            switch (pVal.ItemUID)
            {
                case "0_U_G":
                    DbData.SetValue("LineId", pVal.Row - 1, pVal.Row.ToString());
                    DbData.SetValue("U_MunCode", pVal.Row - 1, MunCode);
                    DbData.SetValue("U_MunName", pVal.Row - 1, MunName);
                    break;
            }
            matrix.LoadFromDataSourceEx();
            matrix.AutoResizeColumns();
            if (oForm.Mode == BoFormMode.fm_OK_MODE) oForm.Mode = BoFormMode.fm_UPDATE_MODE;
        }
        static public void addInsertRowRelationMenuUDO(SAPbouiCOM.Form objForm, SAPbouiCOM.ContextMenuInfo eventInfo)
        {
            MenuCreationParams objParams = null;

            try
            {
                objParams = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objParams.String = "Agregar línea";
                objParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                objParams.UniqueID = "HCO_MWTARU";
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
                if (MainObject.Instance.B1Application.Menus.Exists("HCO_MWTDRU")) MainObject.Instance.B1Application.Menus.RemoveEx("HCO_MWTDRU");
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
        static public void removeInsertRowRelationMenuUDO()
        {
            try
            {
                if (MainObject.Instance.B1Application.Menus.Exists("HCO_MWTARU"))
                {
                    MainObject.Instance.B1Application.Menus.RemoveEx("HCO_MWTARU");
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


        //revisar para optimización
        public static bool activateWTMenu(string FormUID, bool state)
        {
            SAPbouiCOM.MenuItem objMenuItem = null;
            bool activated = false;

            try
            {
                if (!state)
                {
                    Operations.CloseFormBP = true;
                    RelatedParties.Operations.VisibleWith = false;
                    var form = MainObject.Instance.B1Application.Forms.Item(FormUID);

                    RelatedParties.Instance.LoadRelatedPartiesForm(RelatedParties.Settings.RelatedParties.RELATED_PARTIES_DUMMIES, true);
                    var formActiv = MainObject.Instance.B1Application.Forms.ActiveForm;
                    MainObject.Instance.B1Application.Forms.Item(formActiv.UniqueID).Select();
                    formActiv.Close();

                    MainObject.Instance.B1Application.Forms.Item(form.UniqueID).Select();
                }

                objMenuItem = MainObject.Instance.B1Application.Menus.Item("5897");
                if (objMenuItem.Enabled)
                {

                    CacheManager.CacheManager.Instance.addToCache("WTAutoActivate", FormUID, CacheManager.CacheManager.objCachePriority.Default);
                    MainObject.Instance.B1Application.ActivateMenuItem("5897");
                    CacheManager.CacheManager.Instance.removeFromCache("WTAutoActivate");
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

            return activated;
        }
        internal static bool HasRelParty(string CardCode)
        {
            objRS = (SAPbobsCOM.Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            strSQL = string.Format(Queries.Instance.Queries().Get("GetBPCount"), CardCode);

            try
            {
                objRS.DoQuery(strSQL);
                if (objRS.RecordCount > 0)
                {
                    objRS.MoveFirst();
                    return Int32.Parse(objRS.Fields.Item(0).Value.ToString()) > 0;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objRS);
                strSQL = string.Empty;
                GC.Collect();
            }


            return false;
        }

        static public void AddDocumentInfo(BusinessObjectInfo BusinessObjectInfo)
        {
            oXml = new XmlDocument();
            oXml.LoadXml(BusinessObjectInfo.ObjectKey);
            XmlNode Xn = oXml.LastChild;
            BackgroundWorker addDocumentInfoWorker = new BackgroundWorker();
            AddDocumentInfoArgs objArgs = new AddDocumentInfoArgs()
            {
                ObjectKey = Xn["DocEntry"].InnerText,
                ObjectType = BusinessObjectInfo.Type,
                FormtTypeEx = BusinessObjectInfo.FormTypeEx,
                FormUID = BusinessObjectInfo.FormUID
            };
            addDocumentInfoWorker.WorkerSupportsCancellation = false;
            addDocumentInfoWorker.WorkerReportsProgress = false;
            addDocumentInfoWorker.DoWork += AddDocumentInfoWorker_DoWork;
            addDocumentInfoWorker.RunWorkerCompleted += AddDocumentInfoWorker_RunWorkerCompleted;
            addDocumentInfoWorker.RunWorkerAsync(objArgs);
        }
        static private void AddDocumentInfoWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!(e.Error == null)) _Logger.Error(e.Error.Message);
        }
        private static double GetWTDocBaseAmount(SAPbobsCOM.Documents objDoc)
        {
            double dbBase = 0;
            try
            {
                for (int i = 0; i < objDoc.Lines.Count; i++)
                {
                    objDoc.Lines.SetCurrentLine(i);
                    if (objDoc.Lines.WTLiable == BoYesNoEnum.tYES && objDoc.Lines.TaxTotal > 0)
                    {
                        dbBase += objDoc.Lines.LineTotal;
                    }
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                dbBase = -1;
            }
            return dbBase;
        }
        static private void AddDocumentInfoWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            AddDocumentInfoArgs oInfo = null;
            SAPbobsCOM.Documents objDoc = null;
            List<string> WTDocuments = new List<string>();
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralDataParams oFilter = null;
            SAPbobsCOM.GeneralService objEntryObject = null;
            SAPbobsCOM.GeneralData objEntryInfo = null;
            SAPbobsCOM.GeneralData objEntryLinesInfo = null;
            SAPbobsCOM.GeneralDataCollection objEntryLinesObject = null;
            SAPbobsCOM.GeneralService objRelatedPartyObject = null;
            SAPbobsCOM.GeneralData objRelatedPartyInfo = null;
            SAPbobsCOM.BusinessPartners oBP = null;


            string strCardName = String.Empty;
            string strRelatedParty = "";
            string munCode = string.Empty;
            //string dbBaseAmnt = "";

            try
            {
                oInfo = (AddDocumentInfoArgs)e.Argument;
                WTDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._WithHoldingTax.WTFormTypes);

                objDoc = (SAPbobsCOM.Documents)MainObject.Instance.B1Company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), oInfo.ObjectType));

                if (WTDocuments.Contains(oInfo.FormtTypeEx))
                {
                    objDoc.GetByKey(int.Parse(oInfo.ObjectKey));
                    objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                    oBP = (BusinessPartners)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                    oBP.GetByKey(objDoc.CardCode);
                    string add_inv = string.Empty;
                    for (int i = 0; i < oBP.Addresses.Count; i++)
                    {
                        oBP.Addresses.SetCurrentLine(i);
                        if (WTDocuments.Contains(oInfo.FormtTypeEx))
                        {
                            add_inv = objDoc.PayToCode;
                            if (add_inv.Equals(oBP.Addresses.AddressName) && oBP.Addresses.AddressType == BoAddressType.bo_BillTo) munCode = oBP.Addresses.UserFields.Fields.Item("U_HCO_MUNI").Value.ToString();
                        }
                        else if (WTDocuments.Contains(oInfo.FormtTypeEx))
                        {
                            add_inv = objDoc.ShipToCode;
                            if (add_inv.Equals(oBP.Addresses.AddressName) && oBP.Addresses.AddressType == BoAddressType.bo_ShipTo) munCode = oBP.Addresses.UserFields.Fields.Item("U_HCO_MUNI").Value.ToString();
                        }
                    }

                    //dbBaseAmnt = objDoc.WithholdingTaxData.UserFields.Fields.Item("U_HCO_BaseAmnt").Value.ToString();
                    objRelatedPartyObject = objCompanyService.GetGeneralService("HCO_FRP1100");
                    oFilter = (GeneralDataParams)objRelatedPartyObject.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oFilter.SetProperty("Code", GetRelPartyCodeFromCardCode(objDoc.CardCode));

                    objRelatedPartyInfo = objRelatedPartyObject.GetByParams(oFilter);
                    strRelatedParty = objRelatedPartyInfo.GetProperty("Code").ToString();

                    objEntryObject = objCompanyService.GetGeneralService("HCO_FWT1100");
                    objEntryInfo = (GeneralData)objEntryObject.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);
                    objEntryInfo.SetProperty("U_DocNum", objDoc.DocNum);
                    objEntryInfo.SetProperty("U_DocEntry", objDoc.DocEntry);
                    objEntryInfo.SetProperty("U_CardCode", objDoc.CardCode);
                    objEntryInfo.SetProperty("U_RlPyCode", strRelatedParty);
                    objEntryInfo.SetProperty("U_RlPyName", objRelatedPartyInfo.GetProperty("Name"));
                    objEntryInfo.SetProperty("U_DocType", oInfo.ObjectType.Trim());
                    objEntryInfo.SetProperty("U_DocTotal", objDoc.DocTotal);
                    objEntryInfo.SetProperty("U_TransId", objDoc.TransNum);
                    objEntryInfo.SetProperty("U_DocDate", objDoc.DocDate);
                    objEntryInfo.SetProperty("U_OpType", "1");

                    objEntryLinesObject = objEntryInfo.Child("HCO_WT1101");
                    SAPbobsCOM.WithholdingTaxData oWHTData = objDoc.WithholdingTaxData;

                    for (int i = 0; i < oWHTData.Count; i++)
                    {
                        oWHTData.SetCurrentLine(i);
                        SAPbobsCOM.WithholdingTaxCodes oWT = (WithholdingTaxCodes)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oWithholdingTaxCodes);
                        if (oWT.GetByKey(oWHTData.WTCode))
                        {
                            objEntryLinesInfo = objEntryLinesObject.Add();
                            objEntryLinesInfo.SetProperty("U_WTType", oWT.UserFields.Fields.Item("U_HCO_WTType").Value);
                            objEntryLinesInfo.SetProperty("U_WTCode", oWHTData.WTCode);
                            objEntryLinesInfo.SetProperty("U_WTRate", oWT.BaseAmount);
                            objEntryLinesInfo.SetProperty("U_WTBase", objDoc.WithholdingTaxData.UserFields.Fields.Item("U_HCO_BaseAmnt").Value.ToString());
                            objEntryLinesInfo.SetProperty("U_WTAmnt", oWHTData.WTAmount);
                            objEntryLinesInfo.SetProperty("U_BaseLine", oWHTData.LineNum);
                            objEntryLinesInfo.SetProperty("U_Account", oWT.Account);
                            if (oWT.UserFields.Fields.Item("U_HCO_WTType").Value.ToString().Equals("3"))
                            {
                                objEntryLinesInfo.SetProperty("U_MunCode", munCode);
                                objEntryLinesInfo.SetProperty("U_MunName", GetCountyName(munCode));
                            }
                            if (oWT.BaseType == WithholdingTaxCodeBaseTypeEnum.wtcbt_VAT) objEntryLinesInfo.SetProperty("U_WTDocAmnt", GetWTDocBaseAmount(objDoc));
                            //else objEntryLinesInfo.SetProperty("U_WTBase", dbBaseAmnt);
                        }
                    }
                    objEntryObject.Add(objEntryInfo);
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {

            }
        }
        static public void CreateMissingOperations(SAPbouiCOM.ItemEvent pVal)
        {
            List<string> WHPurchaseDocuments = new List<string>();
            List<string> WHSalesDocuments = new List<string>();
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralService objEntryObject = null;
            SAPbobsCOM.GeneralData objEntryInfo = null;
            SAPbobsCOM.GeneralData objEntryLinesInfo = null;
            SAPbobsCOM.GeneralDataParams objResult = null;
            SAPbobsCOM.GeneralDataCollection objEntryLinesObject = null;
            SAPbobsCOM.Documents oDoc = null;
            SAPbobsCOM.GeneralData objRelatedPartyInfo = null;
            SAPbobsCOM.GeneralService objRelatedPartyObject = null;
            SAPbobsCOM.BusinessPartners oBP = null;
            SAPbouiCOM.DataTable DTResult = null;
            SAPbouiCOM.Matrix oMatriz = null;
            Form oForm = MainObject.Instance.B1Application.Forms.ActiveForm;
            string RelPartyCode = string.Empty;
            string dbBaseAmnt = "";
            int count = 0;
            string munCode = string.Empty;


            WHPurchaseDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._WithHoldingTax.WTPurchaseObjectTypes);
            WHSalesDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._WithHoldingTax.WTSalesObjectTypes);
            oMatriz = (Matrix)oForm.Items.Item("grTRA").Specific;
            oMatriz.FlushToDataSource();
            System.Data.DataTable oDT = B1.Base.UIOperations.FormsOperations.SapDataTableToDotNetDataTable(oForm.DataSources.DataTables.Item("DT_TRA").SerializeAsXML(BoDataTableXmlSelect.dxs_All));

            DTResult = oForm.DataSources.DataTables.Item("DT_RES");
            oMatriz = (Matrix)oForm.Items.Item("Item_0").Specific;
            oMatriz.Columns.Item("#").DataBind.Bind("DT_RES", "LineId");
            oMatriz.Columns.Item("Col_0").DataBind.Bind("DT_RES", "DocNum");
            oMatriz.Columns.Item("Col_1").DataBind.Bind("DT_RES", "DocType");
            oMatriz.Columns.Item("Col_2").DataBind.Bind("DT_RES", "Success");
            oMatriz.Columns.Item("Col_3").DataBind.Bind("DT_RES", "Message");

            try
            {
                var selected = from docs in oDT.AsEnumerable()
                               where docs.Field<string>("Seleccionar") == "Y"
                               select new { DocEntry = docs.Field<string>("DocEntry"), DocType = docs.Field<string>("DocType") };

                if (selected.Count() == 0)
                {
                    MainObject.Instance.B1Application.MessageBox("No hay documentos seleccionados para procesar.");
                    return;
                }

                //T1.B1.Base.UIOperations.Operations.startProgressBar("Procesando...", selected.Count());

                foreach (var doc in selected)
                {
                    count++;
                    oDoc = (SAPbobsCOM.Documents)MainObject.Instance.B1Company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), doc.DocType));
                    oBP = (BusinessPartners)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                    oDoc.GetByKey(Int32.Parse(doc.DocEntry));
                    objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                    oBP.GetByKey(oDoc.CardCode);
                    objRelatedPartyObject = objCompanyService.GetGeneralService("HCO_FRP1100");
                    objResult = (GeneralDataParams)objRelatedPartyObject.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                    RelPartyCode = GetRelPartyCodeFromCardCode(oDoc.CardCode);

                    string add_inv = string.Empty;
                    for (int i = 0; i < oBP.Addresses.Count; i++)
                    {
                        oBP.Addresses.SetCurrentLine(i);
                        if (WHPurchaseDocuments.Contains(doc.DocType))
                        {
                            add_inv = oDoc.PayToCode;
                            if (add_inv.Equals(oBP.Addresses.AddressName) && oBP.Addresses.AddressType == BoAddressType.bo_BillTo) munCode = oBP.Addresses.UserFields.Fields.Item("U_HCO_MUNI").Value.ToString();
                        }
                        else if (WHSalesDocuments.Contains(doc.DocType))
                        {
                            add_inv = oDoc.ShipToCode;
                            if (add_inv.Equals(oBP.Addresses.AddressName) && oBP.Addresses.AddressType == BoAddressType.bo_ShipTo) munCode = oBP.Addresses.UserFields.Fields.Item("U_HCO_MUNI").Value.ToString();
                        }
                    }

                    if (!RelPartyCode.Equals(string.Empty))
                    {
                        objResult.SetProperty("Code", RelPartyCode);
                        objRelatedPartyInfo = objRelatedPartyObject.GetByParams(objResult);
                        objEntryObject = objCompanyService.GetGeneralService("HCO_FWT1100");
                        objEntryInfo = (GeneralData)objEntryObject.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);
                        objEntryInfo.SetProperty("U_DocNum", oDoc.DocNum);
                        objEntryInfo.SetProperty("U_DocEntry", oDoc.DocEntry);
                        objEntryInfo.SetProperty("U_CardCode", oDoc.CardCode);
                        objEntryInfo.SetProperty("U_RlPyCode", objRelatedPartyInfo.GetProperty("Code"));
                        objEntryInfo.SetProperty("U_RlPyName", objRelatedPartyInfo.GetProperty("Name"));
                        objEntryInfo.SetProperty("U_DocType", doc.DocType);
                        objEntryInfo.SetProperty("U_DocTotal", oDoc.DocTotal);
                        objEntryInfo.SetProperty("U_TransId", oDoc.TransNum);
                        objEntryInfo.SetProperty("U_DocDate", oDoc.DocDate);
                        objEntryInfo.SetProperty("U_OpType", "1");

                        objEntryLinesObject = objEntryInfo.Child("HCO_WT1101");

                        SAPbobsCOM.WithholdingTaxData oWHTData = oDoc.WithholdingTaxData;
                        for (int i = 0; i < oWHTData.Count; i++)
                        {
                            oWHTData.SetCurrentLine(i);
                            SAPbobsCOM.WithholdingTaxCodes oWT = (WithholdingTaxCodes)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oWithholdingTaxCodes);
                            if (oWT.GetByKey(oWHTData.WTCode))
                            {

                                objEntryLinesInfo = objEntryLinesObject.Add();

                                objEntryLinesInfo.SetProperty("U_WTType", oWT.UserFields.Fields.Item("U_HCO_WTType").Value);
                                objEntryLinesInfo.SetProperty("U_WTCode", oWHTData.WTCode);
                                objEntryLinesInfo.SetProperty("U_WTRate", oWT.BaseAmount);
                                objEntryLinesInfo.SetProperty("U_WTBase", oDoc.WithholdingTaxData.UserFields.Fields.Item("U_HCO_BaseAmnt").Value.ToString());
                                objEntryLinesInfo.SetProperty("U_WTAmnt", oWHTData.WTAmount);
                                objEntryLinesInfo.SetProperty("U_BaseLine", oWHTData.LineNum);
                                objEntryLinesInfo.SetProperty("U_Account", oWT.Account);
                                if (oWT.UserFields.Fields.Item("U_HCO_WTType").Value.ToString().Equals("3"))
                                {
                                    objEntryLinesInfo.SetProperty("U_MunCode", munCode);
                                    objEntryLinesInfo.SetProperty("U_MunName", GetCountyName(munCode));
                                }
                                if (oWT.BaseType == WithholdingTaxCodeBaseTypeEnum.wtcbt_VAT) objEntryLinesInfo.SetProperty("U_WTDocAmnt", GetWTDocBaseAmount(oDoc));
                                //}
                                ////else
                                ////{
                                ////    objEntryLinesInfo.SetProperty("U_WTBase", dbBaseAmnt);
                                ////}
                            }
                        }

                        objEntryObject.Add(objEntryInfo);
                        DTResult.Rows.Add(1);
                        DTResult.SetValue("LineID", count - 1, count);
                        DTResult.SetValue("DocNum", count - 1, oDoc.DocNum);
                        DTResult.SetValue("DocType", count - 1, Int32.Parse(doc.DocType));
                        DTResult.SetValue("Success", count - 1, "Si");
                        DTResult.SetValue("Message", count - 1, "Operación registrada con éxito.");
                        //T1.B1.Base.UIOperations.Operations.setProgressBarMessage(strDocumentType + " " + intDocNum + " procesada.", i + 1);

                    }
                    else
                    {
                        DTResult.Rows.Add(1);
                        DTResult.SetValue("LineID", count - 1, count);
                        DTResult.SetValue("DocNum", count - 1, oDoc.DocNum);
                        DTResult.SetValue("DocType", count - 1, Int32.Parse(doc.DocType));
                        DTResult.SetValue("Success", count - 1, "No");
                        DTResult.SetValue("Message", count - 1, "El Cliente/Proveedor no tiene tercero relacionado");
                    }
                }

                oForm.Freeze(true);
                oMatriz.LoadFromDataSourceEx();
                oMatriz.AutoResizeColumns();
                SAPbouiCOM.Button Obtn = (Button)oForm.Items.Item("btnAdd").Specific;
                Obtn.Caption = "Finalizar";
                oForm.PaneLevel = 2;
                oForm.Freeze(false);
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
                T1.B1.Base.UIOperations.Operations.stopProgressBar();
            }
        }


        static public void RemoveFromCache(string FormId)
        {
            CacheManager.CacheManager.Instance.removeFromCache("Disable_" + FormId);
            CacheManager.CacheManager.Instance.removeFromCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + FormId);
            CacheManager.CacheManager.Instance.removeFromCache(Settings._WithHoldingTax.WTFormInfoCachePrefix + FormId);
            CacheManager.CacheManager.Instance.removeFromCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + FormId);
            CacheManager.CacheManager.Instance.removeFromCache("WTLogicDone_" + FormId);
            CacheManager.CacheManager.Instance.removeFromCache("WTCFLExecuted");
            CacheManager.CacheManager.Instance.removeFromCache("WTLastActiveForm");
        }


    }
}
