using System;
using log4net;
using SAPbouiCOM;

namespace T1.B1.RelatedParties
{
    public class Menu
    {
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static Menu objMenuObject;

        private Menu()
        {
            objMenuObject = new Menu();
        }      

        public static void addThirdPartiesMenu()
        {
            try
            {
                SAPbouiCOM.MenuCreationParams objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "H&&CO Localización";
                objMenu.UniqueID = "HCO_RPT0001";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                int count = MainObject.Instance.B1Application.Menus.Item("8448").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT0000"))
                    MainObject.Instance.B1Application.Menus.Item("8448").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                count = MainObject.Instance.B1Application.Menus.Item("43528").SubMenus.Count + 1;
                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Configuración";
                objMenu.UniqueID = "HCO_MRP0001";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = MainObject.Instance.B1Application.Menus.Item("HCO_RPT0001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MRP0001"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT0001").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Terceros relacionados";
                objMenu.UniqueID = "HCO_MRP0009";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = 0; //MainObject.Instance.B1Application.Menus.Item("43535").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MRP0009"))
                    MainObject.Instance.B1Application.Menus.Item("43535").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Carga de terceros relacionados";
                objMenu.UniqueID = "HCO_MRP0010";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = MainObject.Instance.B1Application.Menus.Item("8704").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MRP0010"))
                    MainObject.Instance.B1Application.Menus.Item("8704").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Reportes Localización";
                objMenu.UniqueID = "HCO_RPT0000";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                count = 0; //MainObject.Instance.B1Application.Menus.Item("43535").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT0000"))
                    MainObject.Instance.B1Application.Menus.Item("43531").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Movimiento por terceros";
                objMenu.UniqueID = "HCO_MRP1010";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = 0; //MainObject.Instance.B1Application.Menus.Item("43535").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MRP1010"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT0000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "ERI (Perdidas y ganancias)";
                objMenu.UniqueID = "HCO_RPT0011";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = 0; //MainObject.Instance.B1Application.Menus.Item("43535").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT0011"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT0000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "ESFA (Balance)";
                objMenu.UniqueID = "HCO_RPT0012";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = 0; //MainObject.Instance.B1Application.Menus.Item("43535").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT0012"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT0000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Balance de prueba";
                objMenu.UniqueID = "HCO_RPT0013";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = 0; //MainObject.Instance.B1Application.Menus.Item("43535").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT0013"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT0000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Movimiento Diario";
                objMenu.UniqueID = "HCO_RPT0014";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = 0; //MainObject.Instance.B1Application.Menus.Item("43535").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT0014"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT0000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Certificado de retención";
                objMenu.UniqueID = "HCO_RPT0015";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = 0; //MainObject.Instance.B1Application.Menus.Item("43535").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT0015"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT0000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Auxiliar por Cuenta";
                objMenu.UniqueID = "HCO_RPT0016";
                objMenu.Type = BoMenuType.mt_STRING;
                count = 0; //MainObject.Instance.B1Application.Menus.Item("43535").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT0016"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT0000").SubMenus.AddEx(objMenu);


                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Retenciones en compras por código";
                objMenu.UniqueID = "HCO_RPT0017";
                objMenu.Type = BoMenuType.mt_STRING;
                count = 0; //MainObject.Instance.B1Application.Menus.Item("43535").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT0018"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT0000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Retenciones en ventas por código";
                objMenu.UniqueID = "HCO_RPT0018";
                objMenu.Type = BoMenuType.mt_STRING;
                count = 0; //MainObject.Instance.B1Application.Menus.Item("43535").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT0019"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT0000").SubMenus.AddEx(objMenu);


                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Retenciones en compras por proveedor";
                objMenu.UniqueID = "HCO_RPT0019";
                objMenu.Type = BoMenuType.mt_STRING;
                count = 0; //MainObject.Instance.B1Application.Menus.Item("43535").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT0020"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT0000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Retenciones en ventas por cliente";
                objMenu.UniqueID = "HCO_RPT0020";
                objMenu.Type = BoMenuType.mt_STRING;
                count = 0; //MainObject.Instance.B1Application.Menus.Item("43535").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT0021"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT0000").SubMenus.AddEx(objMenu);


                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "IVA en ventas por código";
                objMenu.UniqueID = "HCO_RPT0021";
                objMenu.Type = BoMenuType.mt_STRING;
                count = 0; //MainObject.Instance.B1Application.Menus.Item("43535").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT0022"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT0000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "IVA en compras por código";
                objMenu.UniqueID = "HCO_RPT0022";
                objMenu.Type = BoMenuType.mt_STRING;
                count = 0; //MainObject.Instance.B1Application.Menus.Item("43535").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT0023"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT0000").SubMenus.AddEx(objMenu);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }
    }
}
