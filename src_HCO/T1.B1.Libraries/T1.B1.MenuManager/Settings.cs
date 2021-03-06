using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace T1.B1.MenuManager
{
    public class Settings
    {
        public static string AppDataPath { get; set; }
        public static Main _Main { get; }
        public static SelfWithHoldingTax _SelfWithHoldingTax { get; }

        static Settings()
        {

            AppDataPath = AppDomain.CurrentDomain.BaseDirectory + "\\HCO\\T1\\";
            if (!Directory.Exists(AppDataPath))
            {
                Directory.CreateDirectory(Settings.AppDataPath);
            }


            _Main = new Main();
            _Main.Initialize();

            _SelfWithHoldingTax = new SelfWithHoldingTax();
            _SelfWithHoldingTax.Initialize();


        }

        public class Main : Westwind.Utilities.Configuration.AppConfiguration
        {
            public Main()
            {
                logLevel = "Debug";
                mainMenuId = "T1MN001";
                mainMenuDesc = "T1";







            }

            public string logLevel { get; }
            public string mainMenuId { get; }
            public string mainMenuDesc { get; }
            
        }

        public class SelfWithHoldingTax : Westwind.Utilities.Configuration.AppConfiguration
        {
            public SelfWithHoldingTax()
            {
                
                getSelfWithHoldingTaxQuery = "SELECT [@HCO_T1SWT100].\"Code\" ,U_CreditAccount,U_DebitAcct,U_Percent FROM [@HCO_T1SWT100] inner join [@HCO_T1SWT101] on [@HCO_T1SWT100].\"Code\" = [@HCO_T1SWT101].\"Code\" where [@HCO_T1SWT100].U_Enabled = 'Y' and [@HCO_T1SWT101].U_CardCode = '[--CardCode--]'";
                WTaxTransCode = "T1SW";



            }

            
            public string getSelfWithHoldingTaxQuery { get; set; }
            public string WTaxTransCode { get; set; }
        }


    }
}
