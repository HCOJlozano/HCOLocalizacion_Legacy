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

namespace T1.B1.WithholdingTax
{
    public class WithholdingTax
    {
        private static WithholdingTax objWithHoldingTax;
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);

        private WithholdingTax()
        {
            if (objWithHoldingTax == null) objWithHoldingTax = new WithholdingTax();
        }

        static public void getSelectedBPInformation(SAPbouiCOM.ItemEvent pVal, bool useCFL)
        {
            SAPbouiCOM.ChooseFromListEvent oCFLEvent = null;
            SAPbouiCOM.DBDataSource oDB = null;
            SAPbouiCOM.Form objForm = null;
            bool blReadWTConfig = false;
            try
            {
                if (useCFL)
                {
                    oCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                    if (oCFLEvent != null && oCFLEvent.SelectedObjects != null)
                    {
                        bool isDisabled = CacheManager.CacheManager.Instance.getFromCache("Disable_" + pVal.FormUID) == null ? false : true;
                        if (!isDisabled)
                        {
                            string strLastCardCode = CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + pVal.FormUID) == null ? "" : CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + pVal.FormUID);

                            bool blAddOnCalc = CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + pVal.FormUID) == null ? false : true;
                            if (strLastCardCode.Trim().Length == 0)
                            {
                                blAddOnCalc = true;
                            }
                            if (blAddOnCalc)
                            {

                                string strPickedCardCode = oCFLEvent.SelectedObjects.GetValue("CardCode", 0).ToString();
                                if (strLastCardCode.Trim().Length == 0) blReadWTConfig = true;
                                else
                                {
                                    if (strPickedCardCode.Trim() != strLastCardCode) blReadWTConfig = true;
                                }
                                if (blReadWTConfig)
                                {
                                    CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + pVal.FormUID, strPickedCardCode, CacheManager.CacheManager.objCachePriority.Default);
                                    WithholdingTax.getWTforBP(pVal, true);
                                }
                            }
                        }
                    }
                }
                else
                {
                    bool isDisabled = CacheManager.CacheManager.Instance.getFromCache("Disable_" + pVal.FormUID) == null ? false : true;
                    if (!isDisabled)
                    {
                        string strLastCardCode = CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + pVal.FormUID) == null ? "" : CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + pVal.FormUID);

                        bool blAddOnCalc = CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + pVal.FormUID) == null ? false : true;
                        if (strLastCardCode.Trim().Length == 0)
                        {
                            blAddOnCalc = true;
                        }
                        if (blAddOnCalc)
                        {
                            objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                            string strPickedCardCode = string.Empty;

                            switch (objForm.TypeEx)
                            {
                                case "133":
                                    strPickedCardCode = objForm.DataSources.DBDataSources.Item("OINV").GetValue("CardCode", 0).Trim();
                                    break;
                                case "141":
                                    strPickedCardCode = objForm.DataSources.DBDataSources.Item("OPCH").GetValue("CardCode", 0).Trim();
                                    break;
                                case "179":
                                    strPickedCardCode = objForm.DataSources.DBDataSources.Item("ORIN").GetValue("CardCode", 0).Trim();
                                    break;
                                case "181":
                                    strPickedCardCode = objForm.DataSources.DBDataSources.Item("ORPC").GetValue("CardCode", 0).Trim();
                                    break;
                            }

                            if (strLastCardCode.Trim().Length == 0) blReadWTConfig = true;
                            else
                            {
                                if ((strPickedCardCode.Trim() != strLastCardCode) || pVal.ItemUID.Equals("40") || pVal.ItemUID.Equals("10") || pVal.ItemUID.Equals("12") || pVal.ItemUID.Equals("46")) blReadWTConfig = true;
                            }
                            if (blReadWTConfig)
                            {
                                CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + pVal.FormUID, strPickedCardCode, CacheManager.CacheManager.objCachePriority.Default);
                                WithholdingTax.getWTforBP(pVal, useCFL);
                            }
                        }
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
        }
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
        public static void LoadWithHoldingForm(Settings.EWithHoldingTax type)
        {
            try
            {
                FormCreationParams objParams = (FormCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
                objParams.XmlData = GetXmlUDO(type);
                objParams.FormType = GetTypeUDO(type);
                objParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);

                Form objForm = MainObject.Instance.B1Application.Forms.AddEx(objParams);
                objForm.VisibleEx = true;

            }
            catch (Exception er)
            {
                _Logger.Error("(LoadWithHoldingForm)", er);
            }

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
        public static void getWTforBP(SAPbouiCOM.ItemEvent pVal, bool useCFL)
        {
            SAPbouiCOM.Form objForm = null;
            string municipio = string.Empty;
            string strCardCode = "";
            string formType = string.Empty;
            SAPbobsCOM.BusinessPartners objBP = null;
            SAPbobsCOM.WithholdingTaxCodes objWTINfo = null;
            XmlDocument xmlDocument = null;
            XmlNodeList xmlNodes = null;
            XmlNodeList xmlNodesAddr = null;
            XmlNodeList xmlNodesMain = null;
            List<WithholdingTaxConfigDetail> objWithHoldingTaxInfo = null;
            SAPbouiCOM.ChooseFromListEvent oCFLEvent = null;
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                formType = objForm.TypeEx;
                if (useCFL)
                {


                    oCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                    if (oCFLEvent != null)
                    {
                        if (oCFLEvent.SelectedObjects != null)
                        {
                            strCardCode = oCFLEvent.SelectedObjects.GetValue("CardCode", 0).ToString();
                            //strCardCode = (objForm.TypeEx != "133" ? objForm.DataSources.DBDataSources.Item("OPCH").GetValue("CardCode", 0).Trim() : objForm.DataSources.DBDataSources.Item("OINV").GetValue("CardCode", 0).Trim());

                            objBP = (BusinessPartners)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                            if (objBP.GetByKey(strCardCode))
                            {
                                xmlDocument = new XmlDocument();
                                xmlDocument.LoadXml(objBP.GetAsXML());
                                xmlNodes = xmlDocument.SelectNodes("/BOM/BO/BPWithholdingTax/row/WTCode");
                                xmlNodesAddr = xmlDocument.SelectNodes("/BOM/BO/BPAddresses/row");


                                foreach (XmlNode XElement in xmlNodesAddr)
                                {
                                    if (XElement["AddressType"].InnerText.Equals(formType.Equals("133") ? "bo_ShipTo" : "bo_BillTo") && XElement["AddressName"].InnerText.Equals(formType.Equals("133") ? objBP.ShipToDefault : objBP.BilltoDefault))
                                        try
                                        {
                                            municipio = XElement["U_HCO_MUNI"].InnerText;
                                        }
                                        catch
                                        {
                                            municipio = "";
                                        }
                                }

                                if (xmlNodes != null)
                                {
                                    objWithHoldingTaxInfo = new List<WithholdingTaxConfigDetail>();
                                    T1.B1.Base.UIOperations.Operations.startProgressBar("Cargando retenciones asignadas al Socio de negocio", 2);
                                    foreach (XmlNode xn in xmlNodes)
                                    {
                                        string WTCode = xn.InnerText;
                                        objWTINfo = (WithholdingTaxCodes)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oWithholdingTaxCodes);
                                        if (objWTINfo.GetByKey(WTCode))
                                        {
                                            WithholdingTaxConfigDetail objDet = new WithholdingTaxConfigDetail();
                                            objDet.WTCode = WTCode;
                                            objDet.HCO_MMCode = objWTINfo.UserFields.Fields.Item("U_HCO_MMCode").Value.ToString();
                                            objDet.HCO_Area = objWTINfo.UserFields.Fields.Item("U_HCO_Area").Value.ToString();
                                            objDet.HCO_MinBase = double.Parse(objWTINfo.UserFields.Fields.Item("U_HCO_MinBase").Value.ToString());
                                            objDet.HCO_MunGroup = objWTINfo.UserFields.Fields.Item("U_HCO_MunGroup").Value.ToString();
                                            objDet.HCO_WTType = int.Parse(objWTINfo.UserFields.Fields.Item("U_HCO_WTType").Value.ToString());
                                            objDet.MunGroup = getWTMuniInfo(objDet.HCO_MunGroup);

                                            if (objDet.HCO_MunGroup == null || objDet.HCO_MunGroup.Equals(string.Empty))
                                            {
                                                objWithHoldingTaxInfo.Add(objDet);
                                            }
                                            else
                                            {
                                                foreach (WithholdingTaxConfigMun mun in objDet.MunGroup)
                                                {
                                                    if (mun.MunCode.Equals(municipio)) objWithHoldingTaxInfo.Add(objDet);
                                                }
                                            }


                                        }
                                    }
                                    T1.B1.Base.UIOperations.Operations.stopProgressBar();
                                }



                                CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTFOrmInfoCachePrefix + objForm.UniqueID, objWithHoldingTaxInfo, CacheManager.CacheManager.objCachePriority.Default);
                                CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + objForm.UniqueID, true, CacheManager.CacheManager.objCachePriority.Default);
                                //activateWTMenu(objForm.UniqueID);

                            }
                        }
                    }
                }
                else
                {

                    // strCardCode = oCFLEvent.SelectedObjects.GetValue("CardCode", 0);
                    strCardCode = (objForm.TypeEx != "133" ? objForm.DataSources.DBDataSources.Item("OPCH").GetValue("CardCode", 0).Trim() : objForm.DataSources.DBDataSources.Item("OINV").GetValue("CardCode", 0).Trim());

                    objBP = (BusinessPartners)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                    if (objBP.GetByKey(strCardCode))
                    {
                        xmlDocument = new XmlDocument();
                        xmlDocument.LoadXml(objBP.GetAsXML());
                        xmlNodes = xmlDocument.SelectNodes("/BOM/BO/BPWithholdingTax/row/WTCode");
                        xmlNodesAddr = xmlDocument.SelectNodes("/BOM/BO/BPAddresses/row");

                        var ShipToCode = objForm.DataSources.DBDataSources.Item(0).GetValue(formType.Equals("133") ? "ShipToCode" : "BillToCode", 0).ToString();

                        foreach (XmlNode XElement in xmlNodesAddr)
                        {
                            if (XElement["AddressType"].InnerText.Equals(formType.Equals("133") ? "bo_ShipTo" : "bo_BillTo") && XElement["AddressName"].InnerText.Equals(ShipToCode))
                            {
                                try
                                {
                                    municipio = XElement["U_HCO_MUNI"].InnerText;
                                }
                                catch
                                {
                                    municipio = "";
                                }
                                
                            }
                        }

                        if (xmlNodes != null)
                        {
                            objWithHoldingTaxInfo = new List<WithholdingTaxConfigDetail>();
                            T1.B1.Base.UIOperations.Operations.startProgressBar("Cargando retenciones asignadas al Socio de negocio", 2);
                            foreach (XmlNode xn in xmlNodes)
                            {
                                string WTCode = xn.InnerText;
                                objWTINfo = (WithholdingTaxCodes)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oWithholdingTaxCodes);
                                if (objWTINfo.GetByKey(WTCode))
                                {
                                    WithholdingTaxConfigDetail objDet = new WithholdingTaxConfigDetail();
                                    objDet.WTCode = WTCode;
                                    objDet.HCO_MMCode = objWTINfo.UserFields.Fields.Item("U_HCO_MMCode").Value.ToString();
                                    objDet.HCO_Area = objWTINfo.UserFields.Fields.Item("U_HCO_Area").Value.ToString();
                                    objDet.HCO_MinBase = double.Parse(objWTINfo.UserFields.Fields.Item("U_HCO_MinBase").Value.ToString());
                                    objDet.HCO_MunGroup = objWTINfo.UserFields.Fields.Item("U_HCO_MunGroup").Value.ToString();
                                    objDet.HCO_WTType = int.Parse(objWTINfo.UserFields.Fields.Item("U_HCO_WTType").Value.ToString());
                                    objDet.MunGroup = getWTMuniInfo(objDet.HCO_MunGroup);

                                    if (objDet.HCO_MunGroup == null || objDet.HCO_MunGroup.Equals(string.Empty))
                                    {
                                        objWithHoldingTaxInfo.Add(objDet);
                                    }
                                    else
                                    {
                                        foreach (WithholdingTaxConfigMun mun in objDet.MunGroup)
                                        {
                                            if (mun.MunCode.Equals(municipio)) objWithHoldingTaxInfo.Add(objDet);
                                        }
                                    }
                                }
                            }
                            T1.B1.Base.UIOperations.Operations.stopProgressBar();
                        }
                        CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTFOrmInfoCachePrefix + objForm.UniqueID, objWithHoldingTaxInfo, CacheManager.CacheManager.objCachePriority.Default);
                        CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + objForm.UniqueID, true, CacheManager.CacheManager.objCachePriority.Default);
                        //activateWTMenu(objForm.UniqueID);

                    }

                }
            }
            catch (COMException comEx)
            {
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", comEx.ErrorCode, "::", comEx.Message, "::", comEx.StackTrace })));
                _Logger.Error("", exception);
                T1.B1.Base.UIOperations.Operations.stopProgressBar();

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                T1.B1.Base.UIOperations.Operations.stopProgressBar();
            }
        }
        public static void activateWTMenu(string FormUID, bool state)
        {
            SAPbouiCOM.MenuItem objMenuItem = null;

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
        }
        public static void setBPWT(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.Form objDocument = null;
            List<WithholdingTaxConfigDetail> objWithHoldingTaxInfo = null;
            bool blAutoGenerated = false;
            SAPbouiCOM.Matrix objMatrix = null;
            SAPbouiCOM.EditText objEdit = null;
            string strWTCodeValue = "";
            int intNum = -1;
            string strMuniCOdeInAddress = "";
            bool blFirstTime = false;

            try
            {
                blAutoGenerated = CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + FormUID) != null ? true : false;
                if (blAutoGenerated)
                {
                    objWithHoldingTaxInfo = CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTFOrmInfoCachePrefix + FormUID) != null ? CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTFOrmInfoCachePrefix + FormUID) : new List<WithholdingTaxConfigDetail>();

                }

                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                if (pVal.FormTypeEx.Equals("133") || pVal.FormTypeEx.Equals("141"))
                {
                    if (objForm.DataSources.DBDataSources.Item(pVal.FormTypeEx.Equals("133") ? "OINV" : "OPCH").GetValue("CardCode", 0).Equals("")) return;
                }


                objMatrix = (Matrix)objForm.Items.Item("6").Specific;



                if (objWithHoldingTaxInfo != null && objWithHoldingTaxInfo.Count > 0)
                {

                    objDocument = MainObject.Instance.B1Application.Forms.Item(FormUID);
                    strMuniCOdeInAddress = getMuniFromDocument(objDocument);

                    objMatrix.Clear();
                    objMatrix.AddRow(1, -1);

                    #region Fill Matrix
                    intNum = 1;
                    bool blFirst = true;
                    bool blAddressMuni = false;//There is an Muni in the Bp Address
                    if (strMuniCOdeInAddress.Trim().Length > 0)
                    {
                        blAddressMuni = true;
                    }
                    foreach (WithholdingTaxConfigDetail oDetail in objWithHoldingTaxInfo)
                    {

                        bool blCheckMuni = false;
                        //The WT has Muni specific
                        if (blAddressMuni && oDetail.MunGroup.Count > 0)
                        {
                            blCheckMuni = true;
                        }
                        //Get the first column in the matrix first row
                        objEdit = (EditText)objMatrix.GetCellSpecific("1", intNum);
                        string strWTCode = oDetail.WTCode;
                        bool blFound = false;
                        #region Muni Check
                        if (blCheckMuni)
                        {

                            foreach (WithholdingTaxConfigMun oMun in oDetail.MunGroup)
                            {
                                if (oMun.MunCode == strMuniCOdeInAddress)
                                {
                                    blFound = true;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            blFound = true;
                        }
                        #endregion
                        if (!blFound)
                        {
                            strWTCode = "";
                        }
                        if (strWTCode.Trim().Length > 0)
                        {
                            if (!blFirst)
                            {
                                objMatrix.AddRow(1, -1);
                            }
                            objEdit.Value = strWTCode;
                            blFirst = false;
                            intNum++;
                        }
                    }
                    #endregion

                    if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                    objForm.Close();
                }
                else
                {
                    objMatrix.Clear();
                    objMatrix.AddRow(1, -1);

                    objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    if (MainObject.Instance.B1Application.Forms.ActiveForm.UniqueID == pVal.FormUID)
                    {
                        objForm.Close();
                    }
                }
                CacheManager.CacheManager.Instance.addToCache("WTLogicDone_" + FormUID, true, CacheManager.CacheManager.objCachePriority.NotRemovable);
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
        internal static bool HasRelParty(string CardCode)
        {
            Recordset oRS = (SAPbobsCOM.Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = string.Format(Queries.Instance.Queries().Get("GetBPCount"), CardCode);

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

            return false;
        }
        private static List<WithholdingTaxConfigMun> getWTMuniInfo(string strCode)
        {
            SAPbobsCOM.CompanyService objCompany = null;
            SAPbobsCOM.GeneralService objWTInfoService = null;
            SAPbobsCOM.GeneralDataParams objGetParams = null;
            SAPbobsCOM.GeneralData objGeneralData = null;
            SAPbobsCOM.GeneralDataCollection objMuniINfo = null;
            List<WithholdingTaxConfigMun> objList = null;

            try
            {
                if (strCode.Trim().Length > 0)
                {
                    objCompany = MainObject.Instance.B1Company.GetCompanyService();
                    objWTInfoService = objCompany.GetGeneralService(Settings._WithHoldingTax.WTMuniInfoUDO);
                    objGetParams = (GeneralDataParams)objWTInfoService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                    objGetParams.SetProperty("Code", strCode);
                    objGeneralData = objWTInfoService.GetByParams(objGetParams);
                    objMuniINfo = objGeneralData.Child(Settings._WithHoldingTax.WTMuniInfoChildUDO);
                    if (objMuniINfo.Count > 0)
                    {
                        objList = new List<WithholdingTaxConfigMun>();
                        foreach (GeneralData oDef in objMuniINfo)
                        {
                            try
                            {
                                string strMunCode = oDef.GetProperty("U_MunCode").ToString();
                                if (strMunCode != null && strMunCode.Trim().Length > 0)
                                {
                                    WithholdingTaxConfigMun oMun = new WithholdingTaxConfigMun();
                                    oMun.MunCode = strMunCode;
                                    objList.Add(oMun);


                                }
                            }
                            catch (Exception er)
                            {
                                _Logger.Error("", er);
                            }
                        }
                    }
                }
                else
                {
                    objList = new List<WithholdingTaxConfigMun>();
                }

            }
            catch (COMException comEx)
            {
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", comEx.ErrorCode, "::", comEx.Message, "::", comEx.StackTrace })));
                _Logger.Error("", exception);
                objList = new List<WithholdingTaxConfigMun>();

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                objList = new List<WithholdingTaxConfigMun>();
            }
            return objList;
        }
        private static string getMuniFromDocument(SAPbouiCOM.Form objForm)
        {
            string strCode = string.Empty;
            string strCardCode = string.Empty;
            string strAddressCode = "";
            SAPbobsCOM.BusinessPartners objBP = null;
            SAPbobsCOM.BPAddresses objBpAddress = null;

            try
            {
                switch (objForm.TypeEx)
                {
                    case "133":
                        strCardCode = objForm.DataSources.DBDataSources.Item("OINV").GetValue("CardCode", 0).Trim();
                        strAddressCode = objForm.DataSources.DBDataSources.Item("OINV").GetValue("PayToCode", 0).Trim();
                        break;
                    case "141":
                        strCardCode = objForm.DataSources.DBDataSources.Item("OPCH").GetValue("CardCode", 0).Trim();
                        strAddressCode = objForm.DataSources.DBDataSources.Item("OPCH").GetValue("PayToCode", 0).Trim();
                        break;
                    case "179":
                        strCardCode = objForm.DataSources.DBDataSources.Item("ORIN").GetValue("CardCode", 0).Trim();
                        strAddressCode = objForm.DataSources.DBDataSources.Item("ORIN").GetValue("PayToCode", 0).Trim();
                        break;
                    case "181":
                        strCardCode = objForm.DataSources.DBDataSources.Item("ORPC").GetValue("CardCode", 0).Trim();
                        strAddressCode = objForm.DataSources.DBDataSources.Item("ORPC").GetValue("PayToCode", 0).Trim();
                        break;
                }


                objBP = (BusinessPartners)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                if (objBP.GetByKey(strCardCode))
                {
                    objBpAddress = objBP.Addresses;
                    for (int i = 0; i < objBpAddress.Count; i++)
                    {
                        objBpAddress.SetCurrentLine(i);
                        if (objBpAddress.AddressType == BoAddressType.bo_BillTo && objBpAddress.AddressName == strAddressCode)
                        {
                            strCode = objBpAddress.UserFields.Fields.Item("U_HCO_MUNI").Value.ToString();
                            break;
                        }
                    }

                }
            }
            catch (COMException comEx)
            {
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", comEx.ErrorCode, "::", comEx.Message, "::", comEx.StackTrace })));
                _Logger.Error("", exception);
                MainObject.Instance.B1Application.SetStatusBarMessage("HCO: No se pudo recuperar el municipio de la direccion de pago. Las retenciones no se filtraran por municipio", SAPbouiCOM.BoMessageTime.bmt_Short, true);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("HCO: No se pudo recuperar el municipio de la direccion de pago. Las retenciones no se filtraran por municipio", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            return strCode;
        }
        public static void SetChooseFromListMunMatrix(ItemEvent pVal)
        {
            var oForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
            var DbData = oForm.DataSources.DBDataSources.Item(1);
            var matrix = ((Matrix)oForm.Items.Item(pVal.ItemUID).Specific);
            var MunCode = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Code")[0].ToString();
            var MunName = B1.Base.UIOperations.FormsOperations.ListChoiceListener(pVal, "Name")[0].ToString();
            if (MunCode.Equals(string.Empty))
                return;


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
            if (oForm.Mode == BoFormMode.fm_OK_MODE)
                oForm.Mode = BoFormMode.fm_UPDATE_MODE;
        }
        static public bool formModeAdd(SAPbouiCOM.ItemEvent pVal)
        {
            bool blResult = false;
            try
            {
                SAPbouiCOM.Form objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    blResult = true;
                }
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objForm);
                objForm = null;
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                blResult = false;
            }
            return blResult;
        }
        private static double getWTDocBaseAmount(SAPbobsCOM.Documents objDoc)
        {
            double dbBase = 0;
            try
            {
                for(int i = 0; i < objDoc.Lines.Count; i++)
                {
                    objDoc.Lines.SetCurrentLine(i);
                    if(objDoc.Lines.WTLiable == BoYesNoEnum.tYES && objDoc.Lines.TaxTotal > 0)
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
        static public void addDocumentInfo(AddDocumentInfoArgs BusinessObjectInfo)
        {
            BackgroundWorker addDocumentInfoWorker = new BackgroundWorker();
            addDocumentInfoWorker.WorkerSupportsCancellation = false;
            addDocumentInfoWorker.WorkerReportsProgress = false;
            addDocumentInfoWorker.DoWork += AddDocumentInfoWorker_DoWork;
            addDocumentInfoWorker.RunWorkerCompleted += AddDocumentInfoWorker_RunWorkerCompleted;
            addDocumentInfoWorker.RunWorkerAsync(BusinessObjectInfo);

        }
        internal static void InitMissingOperationsForm()
        {
            string strSQL = "";
            SAPbobsCOM.Recordset objRS = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.DataTable objDT = null;
            SAPbouiCOM.Matrix objMatrix = null;
            SAPbouiCOM.GridColumn objGridColumn = null;
            SAPbouiCOM.EditTextColumn oEditTExt = null;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.ActiveForm;
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
            List<string> WHPurchaseDocuments = new List<string>();
            List<string> WHSalesDocuments = new List<string>();
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
            double dbBaseAmnt = 0;

            try
            {
                oInfo = (AddDocumentInfoArgs)e.Argument;
                WHPurchaseDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._WithHoldingTax.WTPurchaseObjects);
                WHSalesDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._WithHoldingTax.WTSalesObjects);

                objDoc = (SAPbobsCOM.Documents)MainObject.Instance.B1Company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), oInfo.ObjectType));

                if (WHPurchaseDocuments.Contains(oInfo.FormtTypeEx) || WHSalesDocuments.Contains(oInfo.FormtTypeEx))
                {
                    objDoc.GetByKey(int.Parse(oInfo.ObjectKey));
                    objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                    oBP = (BusinessPartners)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                    oBP.GetByKey(objDoc.CardCode);
                    string add_inv = string.Empty;  
                    for (int i = 0; i < oBP.Addresses.Count; i++)
                    {
                        oBP.Addresses.SetCurrentLine(i);
                        if (WHPurchaseDocuments.Contains(oInfo.FormtTypeEx))
                        {
                            add_inv = objDoc.PayToCode;
                            if (add_inv.Equals(oBP.Addresses.AddressName) && oBP.Addresses.AddressType == BoAddressType.bo_BillTo)
                            {
                                munCode = oBP.Addresses.UserFields.Fields.Item("U_HCO_MUNI").Value.ToString();
                            }
                        }
                        else if (WHSalesDocuments.Contains(oInfo.FormtTypeEx))
                        {
                            add_inv = objDoc.ShipToCode;
                            if (add_inv.Equals(oBP.Addresses.AddressName) && oBP.Addresses.AddressType == BoAddressType.bo_ShipTo)
                            {
                                munCode = oBP.Addresses.UserFields.Fields.Item("U_HCO_MUNI").Value.ToString();
                            }
                        }
                    }                    

                    dbBaseAmnt = getWTDocBaseAmount(objDoc);
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
                            objEntryLinesInfo.SetProperty("U_WTBase", oWHTData.TaxableAmount);
                            objEntryLinesInfo.SetProperty("U_WTAmnt", oWHTData.WTAmount);
                            objEntryLinesInfo.SetProperty("U_BaseLine", oWHTData.LineNum);
                            objEntryLinesInfo.SetProperty("U_Account", oWT.Account);
                            if (oWT.UserFields.Fields.Item("U_HCO_WTType").Value.ToString().Equals("3"))
                            {
                                objEntryLinesInfo.SetProperty("U_MunCode", munCode);
                                objEntryLinesInfo.SetProperty("U_MunName", GetCountyName(munCode));
                            }
                            if (oWT.BaseType == WithholdingTaxCodeBaseTypeEnum.wtcbt_VAT)
                            {
                                objEntryLinesInfo.SetProperty("U_WTDocAmnt", dbBaseAmnt);
                            }
                            //else
                            //{
                            //    objEntryLinesInfo.SetProperty("U_WTBase", dbBaseAmnt);
                            //}
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

        static public void createMissingOperations(SAPbouiCOM.ItemEvent pVal)
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
            double dbBaseAmnt = 0;
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
                    dbBaseAmnt = getWTDocBaseAmount(oDoc);
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
                            if (add_inv.Equals(oBP.Addresses.AddressName) && oBP.Addresses.AddressType == BoAddressType.bo_BillTo)
                            {
                                munCode = oBP.Addresses.UserFields.Fields.Item("U_HCO_MUNI").Value.ToString();
                            }
                        }
                        else if (WHSalesDocuments.Contains(doc.DocType))
                        {
                            add_inv = oDoc.ShipToCode;
                            if (add_inv.Equals(oBP.Addresses.AddressName) && oBP.Addresses.AddressType == BoAddressType.bo_ShipTo)
                            {
                                munCode = oBP.Addresses.UserFields.Fields.Item("U_HCO_MUNI").Value.ToString();
                            }
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
                                objEntryLinesInfo.SetProperty("U_WTBase", oWHTData.TaxableAmount);
                                objEntryLinesInfo.SetProperty("U_WTAmnt", oWHTData.WTAmount);
                                objEntryLinesInfo.SetProperty("U_BaseLine", oWHTData.LineNum);
                                objEntryLinesInfo.SetProperty("U_Account", oWT.Account);
                                if (oWT.UserFields.Fields.Item("U_HCO_WTType").Value.ToString().Equals("3"))
                                {
                                    objEntryLinesInfo.SetProperty("U_MunCode", munCode);
                                    objEntryLinesInfo.SetProperty("U_MunName", GetCountyName(munCode));
                                }
                                if (oWT.BaseType == WithholdingTaxCodeBaseTypeEnum.wtcbt_VAT)
                                {
                                    objEntryLinesInfo.SetProperty("U_WTDocAmnt", dbBaseAmnt);
                                }
                                //else
                                //{
                                //    objEntryLinesInfo.SetProperty("U_WTBase", dbBaseAmnt);
                                //}
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
        public static double getPercentFromCode(string WTCode)
        {
            double dblPrctBsAmnt = -1;
            SAPbobsCOM.WithholdingTaxCodes objWT = null;
            try
            {
                objWT = (WithholdingTaxCodes)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oWithholdingTaxCodes);
                if (objWT.GetByKey(WTCode))
                {
                    dblPrctBsAmnt = objWT.BaseAmount;
                }
                else
                {
                    dblPrctBsAmnt = -1;
                    _Logger.Error("Could not get Percent info");
                }

            }
            catch (COMException cOMException1)
            {
                _Logger.Error("", cOMException1);
                dblPrctBsAmnt = -1;
            }
            catch (Exception exception2)
            {
                _Logger.Error("", exception2);
                dblPrctBsAmnt = -1;
            }
            return dblPrctBsAmnt;
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
        static public void MatrixOperationUDO(string Action, string ItemId)
        {
            try
            {
                var objForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                var objMatrix = (Matrix)objForm.Items.Item(ItemId).Specific;

                switch (Action)
                {
                    case "Add":
                        //  objForm.DataSources.DBDataSources.Item(1).InsertRecord(objForm.DataSources.DBDataSources.Item(1).Size);
                        //objForm.DataSources.DBDataSources.Item(1).Offset = objForm.DataSources.DBDataSources.Item(1).Size - 1;
                        objMatrix.AddRow(1);
                        objMatrix.FlushToDataSource();
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
        public static void rowNumber(string ItemId)
        {
            var form = MainObject.Instance.B1Application.Forms.ActiveForm;
            var matrix = (Matrix)form.Items.Item(ItemId).Specific;
            for (int i = 1; i <= matrix.RowCount; i++)
                ((EditText)matrix.GetCellSpecific("#", i)).Value = i.ToString();
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
                    if (objRecordSet != null && objRecordSet.RecordCount > 0)
                    {
                        _relPartyCode = objRecordSet.Fields.Item(0).Value.ToString();

                    }
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

    }
}
