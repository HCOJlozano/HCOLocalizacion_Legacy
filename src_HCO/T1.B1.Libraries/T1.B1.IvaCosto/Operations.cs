using System;
using System.Collections.Generic;
using log4net;
using Newtonsoft.Json;
using System.Runtime.InteropServices;
using SAPbouiCOM;
using System.Resources;
using System.Reflection;
using System.Xml;

namespace T1.B1.IvaCosto
{
    public class Operations
    {
        private static Operations objWIvaCosto;
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static bool runResizelogic = true;
        private static List<string> ICPurchaseDocuments = new List<string>();

        private Operations()
        {
            try
            {
                ICPurchaseDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._IvaCosto.ICPurchaseObjects);
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
        public static void formDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool blBubbleEvent)
        {
            XmlDocument oXml = new XmlDocument();
            if (objWIvaCosto == null) objWIvaCosto = new Operations();

            try
            {
                if (!BusinessObjectInfo.BeforeAction)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            if (BusinessObjectInfo.ActionSuccess && ICPurchaseDocuments.Contains(BusinessObjectInfo.FormTypeEx))
                            {
                                AddDocumentInfoArgs objArgs = new AddDocumentInfoArgs();
                                oXml.LoadXml(BusinessObjectInfo.ObjectKey);
                                XmlNode Xn = oXml.LastChild;
                                objArgs.ObjectKey = Xn["DocEntry"].InnerText;
                                objArgs.ObjectType = BusinessObjectInfo.Type;
                                objArgs.FormtTypeEx = BusinessObjectInfo.FormTypeEx;
                                objArgs.FormUID = BusinessObjectInfo.FormUID;
                                IvaCosto.addDocumentInfo(objArgs);
                            }
                            else if (BusinessObjectInfo.ActionSuccess && (BusinessObjectInfo.FormTypeEx == "133" || BusinessObjectInfo.FormTypeEx == "141"))
                            {
                                oXml.LoadXml(BusinessObjectInfo.ObjectKey);
                                IvaCosto.CreateIVAJournal(BusinessObjectInfo);
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

        public static void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {

                if (!pVal.BeforeAction)
                {
                    switch (pVal.EventType)
                    {
                        case BoEventTypes.et_FORM_LOAD:
                            if (pVal.FormTypeEx.Equals("60712"))
                            {
                                XmlDocument oXML = new XmlDocument();
                                oXML.Load(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Forms\\IvaCosto.srf");
                                var oXml = oXML.InnerXml;
                                MainObject.Instance.B1Application.LoadBatchActions(string.Format(oXml, pVal.FormUID));
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
            finally
            {
            }
        }

    }
}
