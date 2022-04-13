using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Westwind.Utilities.Configuration;

namespace T1.B1.IvaCosto
{
    public class Settings
    {
        public static Main _Main { get; }
        public static string AppDataPath { get; set; }
        public static IvaCosto _IvaCosto { get; }
        public enum EIvaCosto
        {
            IVA
        };
        static Settings()
        {
            AppDataPath = AppDomain.CurrentDomain.BaseDirectory + T1.B1.Base.InstallInfo.InstallInfo.Config.configurationBaseFolder;
            if (!Directory.Exists(AppDataPath))
            {
                Directory.CreateDirectory(Settings.AppDataPath);
            }

            _Main = new Main();
            _Main.Initialize();

            _IvaCosto = new IvaCosto();
            _IvaCosto.Initialize();
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

        public class IvaCosto : Westwind.Utilities.Configuration.AppConfiguration
        {
            public string ICFOrmInfoCachePrefix { get; }
            public string ICPurchaseObjects { get; set; }
            public string ICLastCardCodeCachePrefix { get; set; }
            public string ICInternalRegistryUDO { get; set; }
            public IvaCosto()
            {
                ICFOrmInfoCachePrefix = "WTInfoForm_";
                ICLastCardCodeCachePrefix = "WTLastCardCode_";
                ICPurchaseObjects = "[\"143\",\"182\",\"141\",\"181\"]";
                ICInternalRegistryUDO = "HCO_FWT0100";
            }

            protected override Westwind.Utilities.Configuration.IConfigurationProvider OnCreateDefaultProvider(string sectionName, object configData)
            {
                var provider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<IvaCosto>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = provider;

                return provider;
            }
        }
    }
}
