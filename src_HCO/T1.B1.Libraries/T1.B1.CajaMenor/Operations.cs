using System;
using log4net;
using System.Runtime.InteropServices;

namespace T1.B1.CajaMenor
{
    public class Operations
    {
        private static Operations objCajaMenor;
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._MainPettyCash.logLevel);
        

        private Operations()
        {

        }

        
        public static void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
       
            if(objCajaMenor == null) objCajaMenor = new Operations(); 
            try
            {

                if (!pVal.BeforeAction)
                {
                    switch (pVal.MenuUID)
                    {
                        case "HCO_MCM0002":

                                T1.B1.CajaMenor.Forms.Form1 oConceptForm = new Forms.Form1();

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

        }        
    
    }
}
