using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Westwind.Utilities.Configuration;

namespace T1.B1.SelfWithholdingTax
{
    public class Settings
    {
        public static Main _Main { get; }
        public static string AppDataPath { get; set; }
        public static SelfWithHoldingTax _SelfWithHoldingTax { get; }


        static Settings()
        {
            AppDataPath = AppDomain.CurrentDomain.BaseDirectory + T1.B1.Base.InstallInfo.InstallInfo.Config.configurationBaseFolder;
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

        public class SelfWithHoldingTax : Westwind.Utilities.Configuration.AppConfiguration
        {
            public string getMissingSWT { get; }

            public bool useVersion2 { get; }
            public string getAppliedSWTinDOc = "";
            public string lastFolderId { get; }
            public int SelfWithHoldingFolderPane { get; }
            public string SelfWithHoldingFolderId { get; }
            public string showFolderInDocumentsList { get; }
            public string getSelfWithHoldingTaxQuery { get; set; }
            public string getSelfWithHoldingTaxQueryPurchase { get; set; }
            public string WTaxTransCode { get; set; }
            public string CancelFormUID { get; }
            public string getPostedSWtaxQueryV1 { get; }
            public string getPostedSWtaxQueryV2 { get; }
            public bool TransactionCodeBase { get; }
            public string getRegistrationFromJEQuery { get; }
            public string SWtaxUDOTransaction { get; set; }
            public string getSelfWithHoldingTransactions { get; set; }
            public string relatedpartyFieldInLines { get; }
            public string MissingSWTFormUID { get; }
            public string WTSalesObjects { get; set; }
            public string WTPurchaseObjects { get; set; }
            public SelfWithHoldingTax()
            {

                getSelfWithHoldingTaxQuery = "SELECT DISTINCT \"U_MinAmnt\",TA.\"Code\" ,TA.\"U_CreditAcct\", TA.\"U_DebitAcct\", TA.\"U_Rate\" FROM \"@HCO_SW0100\" TA left join \"@HCO_SW0101\" TB on TA.\"Code\" = TB.\"Code\" WHERE (TA.\"U_Sales\" = '{1}') and ((TA.\"U_Enabled\" = 'Y' and TB.\"U_CardCode\" = '{0}') or TA.\"U_IsGlobal\" = 'Y')";
                getSelfWithHoldingTaxQueryPurchase = "SELECT DISTINCT TA.\"U_MinAmnt\", TA.\"Code\" , TA.\"U_CreditAcct\", TA.\"U_DebitAcct\", TA.\"U_Rate\" FROM \"@HCO_SW0100\" TA left join \"@HCO_SW0101\" TB on TA.\"Code\" = TB.\"Code\" WHERE (TA.\"U_Purchase\" = '{1}') and ((TA.\"U_Enabled\" = 'Y' and TB.\"U_CardCode\" = '{0}') or [TA.\"U_IsGlobal\" = 'Y')";

                WTaxTransCode = "T1SW";
                relatedpartyFieldInLines = "U_HCO_RELPAR";
                SWtaxUDOTransaction = "HCO_FSW1100";
                CancelFormUID = "HCO_SWTF001";
                getPostedSWtaxQueryV1 = "select distinct 'N' as \"Sel.\", \"TransId\" as \"Asiento\", \"TaxDate\" as \"Fecha\", case when \"credit\" > 0 then \"Credit\" else \"Debit\" end as \"Total\", \"LineMemo\" as \"Comentario\", space(500) as \"Resultado\" from JDT1 where \"LineMemo\" like '%Auto%[--SWTCode--]%' and(TaxDate >= Convert(datetime, '[--StartDate--]', 112) and \"TaxDate\" <= Convert(datetime, '[--EndDate--]', 112))";
                getPostedSWtaxQueryV2 = "select distinct 'N' as \"Sel.\",OJDT.\"TransId\" as \"Asiento\", OJDT.\"TaxDate\" as \"Fecha\", case when credit > 0 then \"Credit\" else \"Debit\" end as \"Total\", \"Memo\" as \"Comentario\", space(500) as \"Resultado\" " +
" from JDT1 "+
" inner join OJDT on OJDT.\"TransId\" = JDT1.\"TransId\" " +
" where OJDT.\"TransCode\" = '[--TransCode--]' and(OJDT.\"TaxDate\" >= Convert(datetime, '[--StartDate--]', 112) and OJDT.\"TaxDate\" <= Convert(datetime, '[--EndDate--]', 112)) " +
" and OJDT.\"StornoToTr\" is null " +
" and(JDT1.\"TransId\" not in (select distinct \"StornoToTr\" from OJDT where \"StornoToTr\" is not null))";
                    
                TransactionCodeBase = true;
                getRegistrationFromJEQuery = "SELECT \"DocEntry\" from [@HCO_SW1100] where \"U_TrasnId\" = [--JE--]";         
                getSelfWithHoldingTransactions = "select \"U_DocEntry\", \"U_DocNum\" from [@HCO_SW1100] where \"Canceled\" = 'N' and \"U_DocDate\" >= convert(datetime, '[--StartDate--]', 112) and \"U_DocDate\" <= convert(datetime, '[--EndDate--]', 112)";
                showFolderInDocumentsList = "133,179,-133,-179";
                SelfWithHoldingFolderId = "HCO_FSWT";
                SelfWithHoldingFolderPane = 20;
                lastFolderId = "1320002137";
                getAppliedSWTinDOc = "select \"U_SWTCode\" as \"Código\", \"U_BaseAmnt\" as \"Base\", \"U_Total\" as \"Retención\", \"U_TransId\" as \"Asiento\", \"Canceled\" as \"Canceladas\" from \"@HCO_SW1100\" where \"U_DocType\"='{0}' and \"U_DocEntry\" ='{1}'";
                MissingSWTFormUID = "HCO_FSW1200";
                useVersion2 = true;
                getMissingSWT = "select distinct 'N' as \"Sel\", \"DocEntry\" as \"Documento\", \"DocNum\" as \"Número\", \"DocDate\" as \"Fecha\", \"DocTotal\" as \"Total\",space(500) as \"Resultado\" from OINV where \"DocEntry\" not in (select distinct \"U_DocEntry\" from [@HCO_SW1100]) and \"DocDate\" >= convert(datetime, '[--StartDate--]', 112) and \"DocDate\" <= convert(datetime, '[--EndDate--]', 112)";
                WTPurchaseObjects = "[\"141\",\"181\",\"65306\",\"60092\"]";
                WTSalesObjects = "[\"133\",\"179\",\"65303\",\"65307\",\"60091\"]";
            }


            

            protected override Westwind.Utilities.Configuration.IConfigurationProvider OnCreateDefaultProvider(string sectionName, object configData)
            {
                var provider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<SelfWithHoldingTax>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = provider;

                return provider;
            }




        }

      




    }
}
