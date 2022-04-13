using System;
using log4net;
using SAPbouiCOM;

namespace T1.B1.InformesTerceros
{
    public class Menu
    {
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static Menu objMenuObject;

        private Menu()
        {
            objMenuObject = new Menu();
        }


        public static void addITRMenu()
        {
            try
            {
                SAPbouiCOM.MenuCreationParams objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Informes Terceros";
                objMenu.UniqueID = "HCO_MITR01";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                int count = MainObject.Instance.B1Application.Menus.Item("HCO_M001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MITR01"))
                {
                    MainObject.Instance.B1Application.Menus.Item("HCO_M001").SubMenus.AddEx(objMenu);
                }

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Balance por Terceros";
                objMenu.UniqueID = "HCO_MITR02";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = MainObject.Instance.B1Application.Menus.Item("HCO_MITR01").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MITR02"))
                {
                    MainObject.Instance.B1Application.Menus.Item("HCO_MITR01").SubMenus.AddEx(objMenu);
                }

                

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }
    }
}
