using System;
using System.IO;
using System.Runtime.InteropServices;


namespace T1.B1.RelatedParties
{
    public class Settings
    {
        public static Main _Main { get; }
        public static string AppDataPath { get; set; }

        public readonly string BPFormTypeEx = "134";

        public enum RelatedParties
        {
            RELATED_PARTIES,
            RELATED_PARTIES_CONFIGURATION,
            RELATED_PARTIES_TYPES,
            RELATED_PARTIES_DOCUMENT_TYPES,
            RELATED_PARTIES_DEPARTMENT,
            RELATED_PARTIES_MUNICIPALITY,
            RELATED_PARTIES_TYPES_CONTRIB,
            RELATED_PARTIES_ECONOMIC_ACTIVITY,
            RELATED_PARTIES_TRIBUTARY_REGIMEN,
            RELATED_PARTIES_MOVEMENT,
            RELATED_PARTIES_MOVEMENT_DETAILS,
            RELATED_PARTIES_CREATION_WIZARD,
            RELATED_PARTIES_DUMMIES
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
        }

        public class Main : Westwind.Utilities.Configuration.AppConfiguration
        {
            public Main()
            {
                logLevel = "Debug";
                lastFolderId = "1320002081";
                RelatedPartiesFolderId = "HCO_FLRP";
                BPFormTypeEx = "134";
                OutgoingPaymentFormTypeEx = "426";
                ReceiptPaymentFormTypeEx = "170";
                JournalFormTypeEx = "392";
                RelatedPartiesFolderPane = 20;
                EmptyDSMainThirdparties = "EmptyDSMainThirdparties";
                EmptyDSRelationThirdPartied = "EmptyDSRelationThirdPartied";
                BPFormMatrixId = "HCO_I47";
                BPFormHCOEditTextItems = "HCO_I51,HCO_I5,HCO_I7,HCO_I35,HCO_I15,HCO_I17,HCO_I19,HCO_I21,HCO_I31,HCO_I33,HCO_I29,HCO_I41,HCO_I11,HCO_I13,HCO_I39,HCO_I43,HCO_I37,HCO_I23,HCO_I25,HCO_I27";
                lastRightClickEventInfo = "lastRightClickEventInfoBP";
                RelatedPartiesUDO = "HCO_FRP1100";
                RelatedPartiesMovementReport = "HCO_T1RPA500UDO";
            }

            public string logLevel { get; }
            public string lastFolderId { get; }
            public string RelatedPartiesFolderId { get; }
            public string BPFormTypeEx { get; }
            public string OutgoingPaymentFormTypeEx { get; }
            public string ReceiptPaymentFormTypeEx { get; }
            public string JournalFormTypeEx { get; }
            public int RelatedPartiesFolderPane { get; }
            public string EmptyDSMainThirdparties { get; }
            public string EmptyDSRelationThirdPartied { get; }
            public string BPFormHCOEditTextItems { get; set; }
            public string BPFormMatrixId { get; }
            public string lastRightClickEventInfo { get; }
            public string RelatedPartiesUDO { get; }
            public string RelatedPartiesMovementReport { get; }

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
    }
}
