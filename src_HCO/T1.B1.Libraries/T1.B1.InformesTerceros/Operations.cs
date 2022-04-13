using System;
using log4net;
using System.Runtime.InteropServices;

namespace T1.B1.InformesTerceros
{
    public class Operations
    {
        private static Operations InformesTerceros;
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);

        private Operations()
        {

        }


        public static void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            if (InformesTerceros == null)
            {
                InformesTerceros = new Operations();
            }

            BubbleEvent = true;
            try
            {
                if (pVal.MenuUID == "HCO_MITR02"
                    && !pVal.BeforeAction)
                {
                    BalanceTerceros.getTransactionList();
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

    }
}
