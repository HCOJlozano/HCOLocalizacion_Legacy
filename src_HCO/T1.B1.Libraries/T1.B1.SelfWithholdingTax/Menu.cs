using System;
using log4net;
using SAPbouiCOM;

namespace T1.B1.SelfWithholdingTax
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
                objMenu.String = "Autorretenciones";
                objMenu.UniqueID = "HCO_MSW0001";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                count = MainObject.Instance.B1Application.Menus.Item("1536").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MSW0001"))
                {
                    MainObject.Instance.B1Application.Menus.Item("1536").SubMenus.AddEx(objMenu);
                }

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Configuración de autorretenciones";
                objMenu.UniqueID = "HCO_MSW0002";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = MainObject.Instance.B1Application.Menus.Item("15616").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MSW0002"))
                {
                    MainObject.Instance.B1Application.Menus.Item("15616").SubMenus.AddEx(objMenu);
                }

                //objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                //objMenu.String = "Autorretenciones faltantes";
                //objMenu.UniqueID = "HCO_MSW0003";
                //objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                //count = MainObject.Instance.B1Application.Menus.Item("HCO_MSW0001").SubMenus.Count + 1;
                //objMenu.Position = count;
                //if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MSW0003"))
                //{
                //    MainObject.Instance.B1Application.Menus.Item("HCO_MSW0001").SubMenus.AddEx(objMenu);
                //}

                //objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                //objMenu.String = "Cancelar Autoretenciones";
                //objMenu.UniqueID = "HCO_MSW0004";
                //objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                //count = MainObject.Instance.B1Application.Menus.Item("HCO_MSW0001").SubMenus.Count + 1;
                //objMenu.Position = count;
                //if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MSW0004"))
                //{
                //    MainObject.Instance.B1Application.Menus.Item("HCO_MSW0001").SubMenus.AddEx(objMenu);
                //}     
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }
    }
}
