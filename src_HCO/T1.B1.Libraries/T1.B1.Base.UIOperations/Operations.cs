using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using log4net;
using System.Runtime.InteropServices;

namespace T1.B1.Base.UIOperations
{
    public class Operations
    {
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        //private static Operations objOper = null;

        private Operations()
        {

        }

        public static void setStatusBarMessage(string strMessage, bool isError, SAPbouiCOM.BoMessageTime msgTime)
        {
            try
            {
                MainObject.Instance.B1Application.SetStatusBarMessage(strMessage, msgTime, isError);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void startProgressBar(string Message, int Max)
        {
            try
            {
                SAPbouiCOM.ProgressBar objProgressbar = null;
                objProgressbar = T1.CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.progressBarCacheName);
                if (objProgressbar == null)
                {
                    objProgressbar = MainObject.Instance.B1Application.StatusBar.CreateProgressBar(Message, Max, false);
                    objProgressbar.Value = 1;
                    objProgressbar.Text = Message;
                    objProgressbar.Value = 1;
                    CacheManager.CacheManager.Instance.addToCache(T1.CacheManager.Settings._Main.progressBarCacheName, objProgressbar, CacheManager.CacheManager.objCachePriority.Default);
                }
            }
            catch (Exception er)
            {
                CacheManager.CacheManager.Instance.removeFromCache(T1.CacheManager.Settings._Main.progressBarCacheName);
                _Logger.Error("", er);
            }
        }

        public static void stopProgressBar()
        {
            try
            {
                SAPbouiCOM.ProgressBar objProgressbar = null;
                objProgressbar = CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.progressBarCacheName);
                if (objProgressbar != null)
                {
                    objProgressbar.Stop();

                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                CacheManager.CacheManager.Instance.removeFromCache(T1.CacheManager.Settings._Main.progressBarCacheName);
            }
        }

        public static void setProgressBarMessage(string strMessage, int Value)
        {
            try
            {
                SAPbouiCOM.ProgressBar objProgressbar = null;
                objProgressbar = CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.progressBarCacheName);
                if (objProgressbar != null)
                {
                    objProgressbar.Text = strMessage;
                    objProgressbar.Value = Value;

                }
                else
                {
                    startProgressBar(strMessage, Value);
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void toggleSelectCheckBox(SAPbouiCOM.ItemEvent pVal, string dtName, string CellNumber)
        {
            SAPbouiCOM.DataTable variable = null;
            SAPbouiCOM.Form variable1 = null;
            XmlDocument xmlDocument = null;
            XmlNodeList xmlNodeLists = null;
            try
            {
                xmlDocument = new XmlDocument();
                variable1 = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                variable = variable1.DataSources.DataTables.Item(dtName);
                xmlDocument.LoadXml(variable.SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly));
                xmlNodeLists = xmlDocument.SelectNodes("/DataTable/Rows/Row/Cells/Cell[" + CellNumber + "]/Value");
                if (xmlNodeLists.Count > 0)
                {
                    foreach (XmlNode xmlNodes in xmlNodeLists)
                    {
                        if (xmlNodes.InnerText != "Y")
                        {
                            xmlNodes.InnerText = "Y";
                        }
                        else
                        {
                            xmlNodes.InnerText = "N";
                        }
                    }
                }
                variable.LoadSerializedXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly, xmlDocument.InnerXml);
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
                _Logger.Error("", exception2);
            }
        }

        public static SAPbouiCOM.Form openFormfromXML(string strXML, string FormType, bool Modal)
        {
            SAPbouiCOM.Form objForm = null;
            //int intLeft = 0;
            string strModFile = "";
            XmlDocument xmlResult = new XmlDocument();
            string strGUID = "";

            try
            {
                strGUID = Modal ? FormType : Guid.NewGuid().ToString().Substring(1, 10);
                strModFile = strXML.Replace("[--UniqueId--]", strGUID)
                    .Replace("[--FormType--]", FormType);
                MainObject.Instance.B1Application.LoadBatchActions(ref strModFile);
                string strResult = MainObject.Instance.B1Application.GetLastBatchResults();
                xmlResult = new XmlDocument();
                xmlResult.LoadXml(strResult);
                bool errors = xmlResult.SelectSingleNode(Settings._FormLoad.errorPath).HasChildNodes != true ? false : true;
                if (!errors)
                {
                    objForm = MainObject.Instance.B1Application.Forms.Item(strGUID);
                }
            }
            catch (COMException ex)
            {
                _Logger.Error("", ex);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            return objForm;
        }


    }
}
