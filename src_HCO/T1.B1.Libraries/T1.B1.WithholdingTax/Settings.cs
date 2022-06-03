﻿using System;
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
        public static SelfWithHoldingTax _SelfWithHoldingTax { get; }
        public static WithHoldingTax _WithHoldingTax { get; }
        public static HANA _HANA { get; set; }
        public static SQL _SQL { get; set; }

        public enum EWithHoldingTax
        {
            MUNICIPALITY,
            WITHHOLDING_OPERATION,
            MISSING_OPERATIONS
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

            _SelfWithHoldingTax = new SelfWithHoldingTax();
            _SelfWithHoldingTax.Initialize();

            _WithHoldingTax = new WithHoldingTax();
            _WithHoldingTax.Initialize();

            _HANA = new HANA();
            _HANA.Initialize();

            _SQL = new SQL();
            _SQL.Initialize();


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
            public string getWTaxDocuments { get; set; }
            public string getSelfWithHoldingTransactions { get; set; }
            public string relatedpartyFieldInLines { get; }
            public string MissingSWTFormUID { get; }

            public SelfWithHoldingTax()
            {

                getSelfWithHoldingTaxQuery = "SELECT DISTINCT U_MINMOUNT,[@HCO_T1SWT100].\"Code\" ,\"U_CreditAcct\", \"U_DebitAcct\", \"U_Percent\" FROM [@HCO_T1SWT100] left join [@HCO_T1SWT101] on [@HCO_T1SWT100].\"Code\" = [@HCO_T1SWT101].\"Code\" WHERE ([@HCO_T1SWT100].\"U_sales\" = '[--isSales--]') and (([@HCO_T1SWT100].\"U_Enabled\" = 'Y' and [@HCO_T1SWT101].\"U_CardCode\" = '[--CardCode--]') or [@HCO_T1SWT100].\"U_global\" = 'Y')";
                getSelfWithHoldingTaxQueryPurchase = "SELECT DISTINCT U_MINMOUNT,[@HCO_T1SWT100].\"Code\" , \"U_CreditAcct\", \"U_DebitAcct\", \"U_Percent\" FROM [@HCO_T1SWT100] left join [@HCO_T1SWT101] on [@HCO_T1SWT100].\"Code\" = [@HCO_T1SWT101].\"Code\" WHERE ([@HCO_T1SWT100].\"U_purchase\" = '[--isPurchase--]') and (([@HCO_T1SWT100].\"U_Enabled\" = 'Y' and [@HCO_T1SWT101].\"U_CardCode\" = '[--CardCode--]') or [@HCO_T1SWT100].\"U_global\" = 'Y')";

                WTaxTransCode = "T1SW";
                relatedpartyFieldInLines = "U_HCO_RELPAR";
                SWtaxUDOTransaction = "HCO_T1SWT200";
                CancelFormUID = "HCO_SWTF001";
                getPostedSWtaxQueryV1 = "select distinct 'N' as \"Sel.\", \"TransId\" as \"Asiento\", \"TaxDate\" as \"Fecha\", case when \"credit\" > 0 then \"Credit\" else \"Debit\" end as \"Total\", \"LineMemo\" as \"Comentario\", space(500) as \"Resultado\" from JDT1 where \"LineMemo\" like '%Auto%[--SWTCode--]%' and(TaxDate >= Convert(datetime, '[--StartDate--]', 112) and \"TaxDate\" <= Convert(datetime, '[--EndDate--]', 112))";
                getPostedSWtaxQueryV2 = "select distinct 'N' as \"Sel.\",OJDT.\"TransId\" as \"Asiento\", OJDT.\"TaxDate\" as \"Fecha\", case when credit > 0 then \"Credit\" else \"Debit\" end as \"Total\", \"Memo\" as \"Comentario\", space(500) as \"Resultado\" " +
" from JDT1 "+
" inner join OJDT on OJDT.\"TransId\" = JDT1.\"TransId\" " +
" where OJDT.\"TransCode\" = '[--TransCode--]' and(OJDT.\"TaxDate\" >= Convert(datetime, '[--StartDate--]', 112) and OJDT.\"TaxDate\" <= Convert(datetime, '[--EndDate--]', 112)) " +
" and OJDT.\"StornoToTr\" is null " +
" and(JDT1.\"TransId\" not in (select distinct \"StornoToTr\" from OJDT where \"StornoToTr\" is not null))";
                    
                TransactionCodeBase = true;
                getRegistrationFromJEQuery = "SELECT \"DocEntry\" from [@HCO_T1SWT200] where \"U_JEEntry\" = [--JE--]";
                getWTaxDocuments = "T1.B1.WithholdingTax.QRY001";
                getSelfWithHoldingTransactions = "select \"U_DocEntry\", \"U_DocNum\" from [@HCO_T1SWT200] where \"Canceled\" = 'N' and \"U_DocDate\" >= convert(datetime, '[--StartDate--]', 112) and \"U_DocDate\" <= convert(datetime, '[--EndDate--]', 112)";
                showFolderInDocumentsList = "133,179,-133,-179";
                SelfWithHoldingFolderId = "HCO_FSWT";
                SelfWithHoldingFolderPane = 20;
                lastFolderId = "1320002137";
                getAppliedSWTinDOc = "select \"U_SWTCode\" as \"Código\", \"U_BaseAmnt\" as \"Base\", U_Total as \"Retención\", \"U_JEEntry\" as \"Asiento\", \"Canceled\" as \"Canceladas\" from [@HCO_T1SWT200] where \"U_DocType\"='[--DocType--]' and \"U_DocEntry\" ='[--DocEntry--]'";
                MissingSWTFormUID = "HCO_SWTF002";
                useVersion2 = true;
                getMissingSWT = "select distinct 'N' as \"Sel\", \"DocEntry\" as \"Documento\", \"DocNum\" as \"Número\", \"DocDate\" as \"Fecha\", \"DocTotal\" as \"Total\",space(500) as \"Resultado\" from OINV where \"DocEntry\" not in (select distinct \"U_DocEntry\" from [@HCO_T1SWT200]) and \"DocDate\" >= convert(datetime, '[--StartDate--]', 112) and \"DocDate\" <= convert(datetime, '[--EndDate--]', 112)";
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

        public class WithHoldingTax : Westwind.Utilities.Configuration.AppConfiguration
        {
            //public string UDOName { get; set; }
            //public string UDOChiledMuni { get; set; }
            public string WTMuniInfoUDO { get; }
            public string WTMuniInfoChildUDO { get; }
            public string WTFOrmInfoCachePrefix { get; }
            public string WTInfoGenCachePrefix { get; }
            public string WTPurchaseObjects { get; set; }
            public string WTLastCardCodeCachePrefix { get; set; }
            public string WTSalesObjects { get; set; }
            public string WTSalesObjectTypes { get; set; }
            public string WTPurchaseObjectTypes { get; set; }
            public string WTCheckObjects { get; set; }
            public string WTInternalRegistryUDO { get; set; }
            public WithHoldingTax()
            {

                //UDOName = "HCOT1WHT200";
                //UDOChiledMuni = "HCOT1WHT200";
                WTMuniInfoUDO = "HCO_FWT0100";
                WTMuniInfoChildUDO = "HCO_WT0101";
                WTFOrmInfoCachePrefix = "WTInfoForm_";
                WTInfoGenCachePrefix = "WTAddOnGenerated_";
                WTLastCardCodeCachePrefix = "WTLastCardCode_";
                WTPurchaseObjects = "[\"141\",\"181\",\"65306\",\"60092\"]";
                WTSalesObjects = "[\"133\",\"179\",\"65303\",\"65307\",\"60091\"]";
                WTSalesObjectTypes = "[\"13\", \"14\"]";
                WTPurchaseObjectTypes = "[\"18\", \"19\"]";
                WTInternalRegistryUDO = "HCO_FWT0100";


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

        public class SQL : Westwind.Utilities.Configuration.AppConfiguration
        {
            public string getMissingOperationsQuery { get; }
            

            public SQL()
            {
                getMissingOperationsQuery = "select * from ( select \"DocEntry\",\"DocNum\" as \"Numero\",\"CardCode\" as \"Socio de Negocio\",'Factura Proveedor' as \"Documento\", \"DocTotal\" as \"Total del Documento\"  from OPCH where \"DocEntry\" not in " +
                                                "( select distinct \"U_DocEntry\" from [@HCO_T1WHT400] where \"U_DocType\" = 18) and \"ObjType\"=18 and \"DocEntry\" in (Select distinct \"AbsEntry\" from PCH5) " +
                                                " union all "+
" select \"DocEntry\",\"DocNum\" as \"Numero\",\"CardCode\" as \"Socio de Negocio\",'NC Factura Proveedor' as \"Documento\", \"DocTotal\" as \"Total del Documento\" "+
" from ORPC where \"DocEntry\" not in (select distinct \"U_DocEntry\" from [@HCO_T1WHT400] where \"U_DocType\" = 19) and \"ObjType\" = 19 and \"DocEntry\" in (Select distinct \"AbsEntry\" from RPC5) " +
" union all "+
" select \"DocEntry\",\"DocNum\" as \"Numero\",\"CardCode\" as \"Socio de Negocio\",'Factura' as \"Documento\", \"DocTotal\" as \"Total del Documento\" "+
" from OINV where \"DocEntry\" not in (select distinct \"U_DocEntry\" from [@HCO_T1WHT400] where \"U_DocType\" = 13) and \"ObjType\" = 13 and \"DocEntry\" in (Select distinct \"AbsEntry\" from INV5) " +
                                                " union all "+
                                                " select \"DocEntry\",\"DocNum\" as \"Numero\",\"CardCode\" as \"Socio de Negocio\",'NC Factura' as \"Documento\", \"DocTotal\" as \"Total del Documento\" "+
                                                " from ORIN where \"DocEntry\" not in (select distinct \"U_DocEntry\" from [@HCO_T1WHT400] where \"U_DocType\" = 14) and \"ObjType\" = 14 and \"DocEntry\" in (Select distinct \"AbsEntry\" from RIN5) " +
" ) as R order by \"Documento\",\"Numero\" ";

                                                ;
                

            }

            protected override Westwind.Utilities.Configuration.IConfigurationProvider OnCreateDefaultProvider(string sectionName, object configData)
            {
                var provider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<SQL>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = provider;

                return provider;
            }



        }

        public class HANA : Westwind.Utilities.Configuration.AppConfiguration
        {
            public string getMissingOperationsQuery { get; }


            public HANA()
            {
                getMissingOperationsQuery = Properties.Resources.MissingOper;
            }

            protected override Westwind.Utilities.Configuration.IConfigurationProvider OnCreateDefaultProvider(string sectionName, object configData)
            {
                var provider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<HANA>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = provider;

                return provider;
            }



        }




    }
}
