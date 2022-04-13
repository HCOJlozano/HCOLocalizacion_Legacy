using System;
using log4net;
using SAPbouiCOM;

namespace T1.B1.IvaCosto
{
    public class Menu
    {
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static Menu objMenuObject;

        private Menu()
        {
            objMenuObject = new Menu();
        }

        
        public static void addWTMenu()
        {
            try
            {
                SAPbouiCOM.MenuCreationParams objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                int count = MainObject.Instance.B1Application.Menus.Item("1536").SubMenus.Count + 1;

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Transacciones IVA Costo";
                objMenu.UniqueID = "HCO_MIC0001";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = MainObject.Instance.B1Application.Menus.Item("1536").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MWT0001"))
                {
                    MainObject.Instance.B1Application.Menus.Item("1536").SubMenus.AddEx(objMenu);
                }               
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }
    }
}
