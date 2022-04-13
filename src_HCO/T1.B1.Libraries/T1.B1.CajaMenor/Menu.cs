using System;
using System.Text;
using log4net;
using System.Runtime.InteropServices;

namespace T1.B1.CajaMenor
{
    public class Menu
    {
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._MainPettyCash.logLevel);
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


                SAPbouiCOM.MenuCreationParams objMenu = (SAPbouiCOM.MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                int count = MainObject.Instance.B1Application.Menus.Item("1536").SubMenus.Count + 1;

                objMenu = (SAPbouiCOM.MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Caja menor";
                objMenu.UniqueID = "HCO_MCM0001";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                count = MainObject.Instance.B1Application.Menus.Item("1536").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MCM0001"))
                {
                    MainObject.Instance.B1Application.Menus.Item("1536").SubMenus.AddEx(objMenu);
                }

                objMenu = (SAPbouiCOM.MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Conceptos caja menor";
                objMenu.UniqueID = "HCO_MCM0002";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("HCO_MCM0001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MCM0002"))
                {
                    MainObject.Instance.B1Application.Menus.Item("HCO_MCM0001").SubMenus.AddEx(objMenu);
                }

                objMenu = (SAPbouiCOM.MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Apertura caja menor";
                objMenu.UniqueID = "HCO_MCM0003";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("HCO_MCM0001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MCM0003"))
                {
                    MainObject.Instance.B1Application.Menus.Item("HCO_MCM0001").SubMenus.AddEx(objMenu);
                }

                objMenu = (SAPbouiCOM.MenuCreationParams) MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Creación caja menor";
                objMenu.UniqueID = "HCO_MCM0004";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("HCO_MCM0001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MCM0004"))
                {
                    MainObject.Instance.B1Application.Menus.Item("HCO_MCM0001").SubMenus.AddEx(objMenu);
                }

                //objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                //objMenu.String = "Cierre Caja Menor";
                //objMenu.UniqueID = "HCO_MCLM006";
                //objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                //count = MainObject.Instance.B1Application.Menus.Item("HCO_MCL0001").SubMenus.Count + 1;
                //objMenu.Position = count;
                //if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MCLM006"))
                //{
                //    MainObject.Instance.B1Application.Menus.Item("HCO_MCL0001").SubMenus.AddEx(objMenu);
                //}


                objMenu = (SAPbouiCOM.MenuCreationParams)MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Registro gasto caja menor";
                objMenu.UniqueID = "HCO_MCM0005";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("HCO_MCM0001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MCM0005"))
                {
                    MainObject.Instance.B1Application.Menus.Item("HCO_MCM0001").SubMenus.AddEx(objMenu);
                }

                //objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                //objMenu.String = "Arqueo Caja Menor";
                //objMenu.UniqueID = "HCO_MCM005";
                //objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                //count = MainObject.Instance.B1Application.Menus.Item("HCO_MCM001").SubMenus.Count + 1;
                //objMenu.Position = count;
                //if (!MainObject.Instance.B1Application.Menus.Exists("HCO_MCM005"))
                //{
                //    MainObject.Instance.B1Application.Menus.Item("HCO_MCM001").SubMenus.AddEx(objMenu);
                //}



            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }
    }
}
