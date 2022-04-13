using System;
using System.Text;
using log4net;
using System.Runtime.InteropServices;
using SAPbouiCOM;

namespace T1.B1.Expenses
{
    public class Menu
    {
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static Menu objMenuObject;

        private Menu()
        {
            objMenuObject = new Menu();
        }

        public static void addMenu(ref StringBuilder sr)
        {
            try
            {
                addExpensesMenu();
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

        public static void removeMenu(string MenuId)
        {


            try
            {
                MainObject.Instance.B1Application.Menus.RemoveEx(MenuId);


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

        public static void addExpensesMenu()
        {
            try
            {


                SAPbouiCOM.MenuCreationParams objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                int count = MainObject.Instance.B1Application.Menus.Item("1536").SubMenus.Count + 1;

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Legalizaciones";
                objMenu.UniqueID = "HCO_MCLM002";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;


                count = MainObject.Instance.B1Application.Menus.Item("1536").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MCLM002"))
                {
                    MainObject.Instance.B1Application.Menus.Item("1536").SubMenus.AddEx(objMenu);
                }


                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Configuración";
                objMenu.UniqueID = "HCO_MCLM007";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;

                count = MainObject.Instance.B1Application.Menus.Item("HCO_MCLM002").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MCLM007"))
                {
                    MainObject.Instance.B1Application.Menus.Item("HCO_MCLM002").SubMenus.AddEx(objMenu);
                }


                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Clasificación Legalizaciones";
                objMenu.UniqueID = "HCO_MCLM009";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("HCO_MCLM007").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MCLM009"))
                {
                    MainObject.Instance.B1Application.Menus.Item("HCO_MCLM007").SubMenus.AddEx(objMenu);
                }


                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Conceptos";
                objMenu.UniqueID = "HCO_MCLM010";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("HCO_MCLM007").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MCLM010"))
                {
                    MainObject.Instance.B1Application.Menus.Item("HCO_MCLM007").SubMenus.AddEx(objMenu);
                }

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Tipos de Legalizaciones";
                objMenu.UniqueID = "HCO_MCLM011";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("HCO_MCLM007").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MCLM011"))
                {
                    MainObject.Instance.B1Application.Menus.Item("HCO_MCLM007").SubMenus.AddEx(objMenu);
                }

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Solicitud de Legalización";
                objMenu.UniqueID = "HCO_MCLM012";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("HCO_MCLM002").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MCLM012"))
                {
                    MainObject.Instance.B1Application.Menus.Item("HCO_MCLM002").SubMenus.AddEx(objMenu);
                }

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Aprobación de Solicitud";
                objMenu.UniqueID = "HCO_MCLM013";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("HCO_MCLM002").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MCLM013"))
                {
                    MainObject.Instance.B1Application.Menus.Item("HCO_MCLM002").SubMenus.AddEx(objMenu);
                }

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Desembolsos";
                objMenu.UniqueID = "HCO_MCLM014";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("HCO_MCLM002").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MCLM014"))
                {
                    MainObject.Instance.B1Application.Menus.Item("HCO_MCLM002").SubMenus.AddEx(objMenu);
                }

                objMenu = (MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Legalización";
                objMenu.UniqueID = "HCO_MCLM015";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("HCO_MCLM002").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MCLM015"))
                {
                    MainObject.Instance.B1Application.Menus.Item("HCO_MCLM002").SubMenus.AddEx(objMenu);
                }

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }
    }
}
