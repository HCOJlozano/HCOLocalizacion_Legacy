using System;
using System.Text;
using log4net;
using System.Runtime.InteropServices;
using SAPbouiCOM;

namespace T1.B1.MenuManager
{
    public class Operations
    {
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static Operations objMenuManager;
        
        private Operations()
        {
            objMenuManager = new Operations();
        }

        public static void addMenu()
        {
            try
            {
                addMainMenu();

                WithholdingTax.Menu.addWTMenu();
                SelfWithholdingTax.Menu.addWTMenu();
                RelatedParties.Menu.addThirdPartiesMenu();
                Expenses.Menu.addExpensesMenu();
                //InformesTerceros.Menu.addITRMenu();
                //CajaMenor.Menu.addExpensesMenu();
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

        public static void removeMenu()
        {
            try
            {

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

        private static void addMainMenu(ref StringBuilder sbMenu)
        {
            sbMenu.Append(Properties.Resources.Menu);
        }

        private static void addMainMenu()
        {
            string ruta = string.Empty;
            try
            {
                SAPbouiCOM.MenuCreationParams objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "T1";
                ruta = "C:\\SAP\\Original\\T1.png";
                objMenu.Image = ruta;//AppDomain.CurrentDomain.BaseDirectory + "Original\\T1.png";
                objMenu.UniqueID = "HCO_M001";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                int count = MainObject.Instance.B1Application.Menus.Item("43520").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_M001"))
                {
                    MainObject.Instance.B1Application.Menus.Item("43520").SubMenus.AddEx(objMenu);
                }
                SAPbouiCOM.IMenuItem objM = MainObject.Instance.B1Application.Menus.Item("43520");
                string strTest = objM.SubMenus.GetAsXML();
            }
            catch (COMException comEx)
            {
                _Logger.Error("" + ruta, comEx);
            }
            catch (Exception er)
            {
                _Logger.Error("" + ruta, er);
            }
        }

    }
}
