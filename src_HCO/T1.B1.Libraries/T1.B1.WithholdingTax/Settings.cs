using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Westwind.Utilities.Configuration;

namespace T1.B1.WithholdingTax
{
    public class Settings
    {
        public static Main _Main { get; }
        public static string AppDataPath { get; set; }
        public static WithHoldingTax _WithHoldingTax { get; }
        public enum EWithHoldingTax
        {
            MUNICIPALITY,
            WITHHOLDING_OPERATION,
            MISSING_OPERATIONS
        };
        static Settings()
        {
            AppDataPath = AppDomain.CurrentDomain.BaseDirectory + T1.B1.Base.InstallInfo.InstallInfo.Config.configurationBaseFolder;
            if (!Directory.Exists(AppDataPath)) Directory.CreateDirectory(Settings.AppDataPath);

            _Main = new Main();
            _Main.Initialize();

            _WithHoldingTax = new WithHoldingTax();
            _WithHoldingTax.Initialize();

        }

        public class Main : Westwind.Utilities.Configuration.AppConfiguration
        {
            public string logLevel { get; }
            public string lastRightClickEventInfo { get; }

            public Main()
            {
                logLevel = "Debug";
                lastRightClickEventInfo = "lastRightClickEventInfoBP";
            }

            protected override Westwind.Utilities.Configuration.IConfigurationProvider OnCreateDefaultProvider(string sectionName, object configData)
            {
                var provider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<Main>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = provider;

                return provider;
            }
        }

        public class WithHoldingTax : Westwind.Utilities.Configuration.AppConfiguration
        {
            public string WTMuniInfoUDO { get; }
            public string WTMuniInfoChildUDO { get; }
            public string WTInternalRegistryUDO { get; set; }


            public string WTFormInfoCachePrefix { get; }
            public string WTInfoGenCachePrefix { get; }
            public string WTLastCardCodeCachePrefix { get; set; }
            public string WTLineTotalsGenCachePrefix { get; set; }

            public string WTFormTypes { get; set; }
            public string WTObjects { get; set; }
            public string WTSalesObjectTypes { get; set; }
            public string WTPurchaseObjectTypes { get; set; }
            public string WTCheckObjects { get; set; }



            public WithHoldingTax()
            {
                WTMuniInfoUDO = "HCO_FWT0100";
                WTMuniInfoChildUDO = "HCO_WT0101";
                WTInternalRegistryUDO = "HCO_FWT0100";

                WTFormInfoCachePrefix = "WTInfoForm_";
                WTInfoGenCachePrefix = "WTAddOnGenerated_";
                WTLastCardCodeCachePrefix = "WTLastCardCode_";
                WTLineTotalsGenCachePrefix = "WTLineTotals";

                WTFormTypes = "[\"141\",\"181\",\"65306\",\"60092\",\"133\",\"179\",\"65303\",\"65307\",\"60091\"]";
                WTObjects = "[\"13\", \"14\",\"18\", \"19\"]";
            }

            protected override Westwind.Utilities.Configuration.IConfigurationProvider OnCreateDefaultProvider(string sectionName, object configData)
            {
                var provider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<WithHoldingTax>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = provider;

                return provider;
            }

        }

    }
}
