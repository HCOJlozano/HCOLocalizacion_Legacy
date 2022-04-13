using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Westwind.Utilities.Configuration;

namespace T1.B1.CajaMenor
{
    public class Settings
    {
        public static MainPettyCash _MainPettyCash { get; set; }

        public static string AppDataPath { get; set; }
        

        static Settings()
        {
            AppDataPath = AppDomain.CurrentDomain.BaseDirectory + T1.B1.Base.InstallInfo.InstallInfo.Config.configurationBaseFolder;
            if (!Directory.Exists(AppDataPath))
            {
                Directory.CreateDirectory(Settings.AppDataPath);
            }


            _MainPettyCash = new MainPettyCash();
            _MainPettyCash.Initialize();
        }

        public class MainPettyCash : Westwind.Utilities.Configuration.AppConfiguration
        {
            public MainPettyCash()
            {
                logLevel = "Debug";
                pettyCashExpenseType = "CM";
                pettyCashPaymentFormType = "HCO_T1PTC001";
                pettyCashUDO = "HCO_T1PTC100";
                pettyCashLegalizationUDO = "HCO_T1PTC300";
                pettyCashConceptUDO = "HCO_T1PTC200";
                pettyCashConceptUDOFormType = "UDO_FT_HCO_T1PTC200";
                pettyCashLegalizationFormType = "UDO_FT_HCO_T1PTC300";
                PCLegalizationFormLastId = "PCLegalizationFormLastId";
                PCLegalizationTransactionCode = "T1A2";
                SysCurrDeviationAccount = "429581005";
            }

            public string logLevel { get; set; }
            public string pettyCashExpenseType { get; set; }
            public string pettyCashPaymentFormType { get; set; }
            public string pettyCashUDO { get; set; }
            public string pettyCashLegalizationUDO { get; set; }
            public string pettyCashConceptUDO { get; set; }
            public string pettyCashConceptUDOFormType { get; set; }
            public string pettyCashLegalizationFormType { get; set; }
            public string PCLegalizationFormLastId { get; set; }
            public string PCLegalizationTransactionCode { get; set; }
            public string SysCurrDeviationAccount { get; set; }


            protected override Westwind.Utilities.Configuration.IConfigurationProvider OnCreateDefaultProvider(string sectionName, object configData)
            {
                var provider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<MainPettyCash>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = provider;

                return provider;
            }

        }






    }
}
