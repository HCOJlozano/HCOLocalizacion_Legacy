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
                count = 0;
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
                objMenu.String = "Reportes localización";
                objMenu.UniqueID = "HCO_RPT0000";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT0000"))
                    MainObject.Instance.B1Application.Menus.Item("43531").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Contables";
                objMenu.UniqueID = "HCO_RPT1000";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT1000"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT0000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Financieros";
                objMenu.UniqueID = "HCO_RPT2000";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                count = 0; //MainObject.Instance.B1Application.Menus.Item("43535").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT2000"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT0000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Impuestos";
                objMenu.UniqueID = "HCO_RPT3000";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT3000"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT0000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "IVA";
                objMenu.UniqueID = "HCO_RPT3100";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                count = 0; 
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT3100"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT3000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Retenciones";
                objMenu.UniqueID = "HCO_RPT3200";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT3200"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT3000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Certificados";
                objMenu.UniqueID = "HCO_RPT3300";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT3300"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT3000").SubMenus.AddEx(objMenu);


                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Auxiliar por Cuenta";
                objMenu.UniqueID = "HCO_RPT1001";
                objMenu.Type = BoMenuType.mt_STRING;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT1001"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT1000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Movimiento Diario";
                objMenu.UniqueID = "HCO_RPT1002";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT1002"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT1000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Movimiento por terceros";
                objMenu.UniqueID = "HCO_RPT1003";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT1003"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT1000").SubMenus.AddEx(objMenu);


                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "ESF (Balance)";
                objMenu.UniqueID = "HCO_RPT2001";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT2001"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT2000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "ERI (Perdidas y ganancias)";
                objMenu.UniqueID = "HCO_RPT2002";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT2002"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT2000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Balance de prueba";
                objMenu.UniqueID = "HCO_RPT2003";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT2003"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT2000").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Balance de prueba por tercero";
                objMenu.UniqueID = "HCO_RPT2004";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT2004"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT2000").SubMenus.AddEx(objMenu);


                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "IVA en ventas por código";
                objMenu.UniqueID = "HCO_RPT3101";
                objMenu.Type = BoMenuType.mt_STRING;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT3101"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT3100").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "IVA en compras por código";
                objMenu.UniqueID = "HCO_RPT3102";
                objMenu.Type = BoMenuType.mt_STRING;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT3102"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT3100").SubMenus.AddEx(objMenu);


                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Retenciones en compras por código";
                objMenu.UniqueID = "HCO_RPT3201";
                objMenu.Type = BoMenuType.mt_STRING;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT3201"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT3200").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Retenciones en ventas por código";
                objMenu.UniqueID = "HCO_RPT3202";
                objMenu.Type = BoMenuType.mt_STRING;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT3202"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT3200").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Retenciones en compras por proveedor";
                objMenu.UniqueID = "HCO_RPT3203";
                objMenu.Type = BoMenuType.mt_STRING;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT3203"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT3200").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Retenciones en ventas por cliente";
                objMenu.UniqueID = "HCO_RPT3204";
                objMenu.Type = BoMenuType.mt_STRING;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT3204"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT3200").SubMenus.AddEx(objMenu);


                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Certificado de retención";
                objMenu.UniqueID = "HCO_RPT3301";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT3301"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT3300").SubMenus.AddEx(objMenu);

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Certificado de retención IVA";
                objMenu.UniqueID = "HCO_RPT3302";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = 0;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_RPT3302"))
                    MainObject.Instance.B1Application.Menus.Item("HCO_RPT3300").SubMenus.AddEx(objMenu);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }
    }
}
