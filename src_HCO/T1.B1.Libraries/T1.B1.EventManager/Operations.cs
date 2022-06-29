using System;
using System.Collections.Generic;
using log4net;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using SAPbouiCOM;

namespace T1.B1.EventManager
{
    public class Operations
    {
        private SAPbouiCOM.Application objApplication = null;
        public bool objStatus = false;
        private static readonly ILog _Logger = Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);


        public bool Status
        {
            get
            {
                return objStatus;
            }
        }

        public Operations()
        {
            try
            {
                objApplication = MainObject.Instance.B1Application;
                objApplication.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(objApplication_AppEvent);
                objApplication.EventLevel = SAPbouiCOM.BoEventLevelType.elf_GlobalEvent;

                
                objApplication.FormDataEvent += FormDataAddEvent;
                objApplication.MenuEvent += MenuEvent;
                objApplication.ItemEvent += ItemEvent;
                objApplication.RightClickEvent += RightClickEvent;
               

                objStatus = true;

                MainObject.Instance.B1Application = objApplication;
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

        public void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (BubbleEvent)
                WithholdingTax.Operations.ItemEvent(FormUID, ref pVal, ref BubbleEvent);

            if (BubbleEvent)
                RelatedParties.Operations.ItemEvent(FormUID, ref pVal, ref BubbleEvent);

            if (BubbleEvent)
                SelfWithholdingTax.Operations.ItemEvent(FormUID, ref pVal, ref BubbleEvent);

            if (BubbleEvent)
                Expenses.Operations.ItemEvent(FormUID, ref pVal, ref BubbleEvent);

            if (BubbleEvent)
                IvaCosto.Operations.ItemEvent(FormUID, ref pVal, ref BubbleEvent);
        }


        public void FormDataAddEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool blBubbleEvent)
        {
            blBubbleEvent = true;

            if (blBubbleEvent)
                WithholdingTax.Operations.FormDataEvent(ref BusinessObjectInfo, ref blBubbleEvent);

            if (blBubbleEvent)
                RelatedParties.Operations.FormDataAddEvent(ref BusinessObjectInfo, ref blBubbleEvent);

            if (blBubbleEvent)
                SelfWithholdingTax.Operations.formDataEvent(ref BusinessObjectInfo, ref blBubbleEvent);

            if (blBubbleEvent)
                Expenses.Operations.formDataAddEvent(ref BusinessObjectInfo, ref blBubbleEvent);

            if (blBubbleEvent)
                IvaCosto.Operations.formDataEvent(ref BusinessObjectInfo, ref blBubbleEvent);
        }

        public void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (BubbleEvent)
                WithholdingTax.Operations.MenuEvent(ref pVal, ref BubbleEvent);

            if (BubbleEvent)
                RelatedParties.Operations.MenuEvent(ref pVal, ref BubbleEvent);

            if (BubbleEvent)
                SelfWithholdingTax.Operations.MenuEvent(ref pVal, ref BubbleEvent);

            if (BubbleEvent)
                Expenses.Operations.MenuEvent(ref pVal, ref BubbleEvent);

            if (BubbleEvent)
                CajaMenor.Operations.MenuEvent(ref pVal, ref BubbleEvent);

            if (BubbleEvent)
                InformesTerceros.Operations.MenuEvent(ref pVal, out BubbleEvent);
        }

        public void RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (BubbleEvent)
                WithholdingTax.Operations.RightClickEvent(ref eventInfo, ref BubbleEvent);

            if (BubbleEvent)
                SelfWithholdingTax.Operations.RightClickEvent(ref eventInfo, ref BubbleEvent);

            if (BubbleEvent)
                RelatedParties.Operations.RightClickEvent(ref eventInfo, ref BubbleEvent);

            if (BubbleEvent)
                Expenses.Operations.RightClickEvent(ref eventInfo, ref BubbleEvent);
        }

        private void ObjApplication_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            throw new NotImplementedException();
        }

        private void ObjApplication_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;                 

        }

        internal class ItemInfo
        {
            public string WareHouse { get; set; }
            public string Dimension1 { get; set; }
            public string Dimension2 { get; set; }
            public string Dimension3 { get; set; }
            public string Dimension4 { get; set; }
            public string Dimension5 { get; set; }
        }



        void objApplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            try
            {
                switch (EventType)
                {
                    case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:

                        System.Windows.Forms.Application.Exit();
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                        System.Windows.Forms.Application.Exit();
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:


                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:

                        System.Windows.Forms.Application.Exit();
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                        System.Windows.Forms.Application.Exit();
                        break;
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
