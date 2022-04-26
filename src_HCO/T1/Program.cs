using System;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using log4net;
using T1.Structure;
using T1.B1;
using System.Management;
using Microsoft.Win32;

namespace T1
{
    static class Program
    {
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                if(!checkInstalled("SAP Crystal Reports runtime engine for .NET Framework (64-bit)") )
                {
                    InstallDependency();
                }

                _Logger.Debug("Adding ConnectionString to cache");
                T1.CacheManager.CacheManager.Instance.addToCache(T1.CacheManager.Settings._Main.connStringCacheName, (string)Environment.GetCommandLineArgs().GetValue(1), CacheManager.CacheManager.objCachePriority.NotRemovable);

                _Logger.Debug("Starting Connection to SAP Business One");
                T1.B1.Connection.Class objConnClass = new B1.Connection.Class();
                objConnClass.B1Connect(false, args);

                _Logger.Debug("Connection status " + objConnClass.Connected.ToString());
                if (objConnClass.Connected)
                {
                    _Logger.Debug("Starting Meta data creation.");

                    bool blMD = T1.B1.MetaData.Operations.blCreateMD(Settings._Main.createMD);
                    _Logger.Debug("Meta data creation ended.");
                    //if (blMD)
                    //{

                    #region loadFirstTimeData
                    //if (Settings._Main.loadInitialData)
                    //{
                    //Expand the logic of this methods in the future to create also all the transaction codes needed
                    var md = new MetaData(MainObject.Instance.B1Company, AppDomain.CurrentDomain.BaseDirectory + "Structure\\RP_Structure.json");
                    md.CreateStructure();
                    md = new MetaData(MainObject.Instance.B1Company, AppDomain.CurrentDomain.BaseDirectory + "Structure\\WT_Structure.json");
                    md.CreateStructure();
                    md = new MetaData(MainObject.Instance.B1Company, AppDomain.CurrentDomain.BaseDirectory + "Structure\\SW_Structure.json");
                    md.CreateStructure();
                    //md = new MetaData(MainObject.Instance.B1Company, AppDomain.CurrentDomain.BaseDirectory + "Structure\\CM_Structure.json");
                    //md.CreateStructure();
                    md = new MetaData(MainObject.Instance.B1Company, AppDomain.CurrentDomain.BaseDirectory + "Structure\\IC_Structure.json");
                    md.CreateStructure();
                    md = new MetaData(MainObject.Instance.B1Company, AppDomain.CurrentDomain.BaseDirectory + "Structure\\TS_Structure.json");
                    md.CreateStructure();

                    T1.B1.MetaData.Operations.loadMuni();
                    T1.B1.MetaData.Operations.loadDepto();
                    T1.B1.MetaData.Operations.loadTipContrib();
                    T1.B1.MetaData.Operations.loadTipDoc();
                    T1.B1.MetaData.Operations.loadActivEcon();
                    T1.B1.MetaData.Operations.loadGenericUDO();



                    #endregion

                    T1.B1.EventFilter.Operations objEventFilter = new B1.EventFilter.Operations();
                    T1.B1.EventManager.Operations objEventManager = new B1.EventManager.Operations();

                   // T1.TaskScheduler.Instance.TaskScheduler();
                    if (objEventManager.Status)
                    {
                        //Add Logic to menu that finds if the menu was created or not before. Find a quicker way to compare the existing menus instead of com object.
                        T1.B1.MenuManager.Operations.addMenu();
                        GC.KeepAlive(objEventManager);
                        GC.KeepAlive(objEventFilter);
                        Application.Run();
                    }
                    else
                    {
                        _Logger.Error("There was an error adding the EventListeners for the AddOn. Please check the log.");
                        T1.B1.MainObject.Instance.B1Application.SetStatusBarMessage("T1: There was an error adding the EventListeners for the AddOn. The execution will be halted.", SAPbouiCOM.BoMessageTime.bmt_Short);
                        Application.Exit();
                    }
                }
                else
                {
                    _Logger.Error("Could not connect to SAP Business One. T1 is terminating.");
                    Application.Exit();
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static string GetProductCode(string productName)
        {
            string query = string.Format("select * from Win32_Product where Name='{0}'", productName);
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(query))
            {
                foreach (ManagementObject product in searcher.Get())
                    return product["IdentifyingNumber"].ToString();
            }
            return null;
        }

        public static bool checkInstalled(string c_name)
        {
            string registryKey = @"Installer\Products";
            RegistryKey key = Registry.ClassesRoot.OpenSubKey(registryKey);

            try
            {
                var valueKey = key.GetSubKeyNames().Where(keyName => (key.OpenSubKey(keyName).GetValue("ProductName") != null ? key.OpenSubKey(keyName).GetValue("ProductName").ToString() : "") == c_name);
                if (valueKey.Count() > 0)
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        public static void InstallDependency()
        {
            var currentPath = System.IO.Directory.GetCurrentDirectory() + "\\Dependencies\\CR13SP31MSI64_0-10010309.msi";
            var installerProcess = new Process();
            var processInfo = new ProcessStartInfo();
                processInfo.Arguments = $@"/i  {currentPath}";
                processInfo.FileName = "msiexec";

            installerProcess.StartInfo = processInfo;
            installerProcess.Start();
            installerProcess.WaitForExit();
        }

    }
}
