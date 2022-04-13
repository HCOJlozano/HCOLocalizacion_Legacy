using System;
using System.Collections;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

namespace T1.CrystalReport
{
    public partial class Form1 : Form
    {
        private ReportDocument report;
        private string PathRep;
        private string RepToExport;
        private string UsrCode;
        private string DBCompany;
        private string DBServer;
        private string DBUserName;
        private string DBPassword;

        public Form1(string pathRep, string repToExp, string usrCode, string bdCompany, string bdServer, string dbUser, string dbPass)
        {
            PathRep = pathRep;
            RepToExport = repToExp;
            UsrCode = usrCode;
            DBCompany = bdCompany;
            DBServer = bdServer;
            DBUserName = dbUser;
            DBPassword = dbPass;

            InitializeComponent();
        }

        private int ConfigureCrystalReports()
        {
            bool impresionAutomatica = false;
            bool exportar_reporte = false;
            string printerName = "";

            try
            {
                PathRep = PathRep.Replace("%", " ").ToString();
                RepToExport = RepToExport.Replace("%", " ");
                ArrayList arrayList = new ArrayList();
                arrayList.Add(UsrCode);

                try
                {
                    report = new ReportDocument();

                    try
                    {
                        try
                        {
                            report.Load(@PathRep, OpenReportMethod.OpenReportByTempCopy);
                        }
                        catch (ApplicationException f)
                        {
                            MessageBox.Show(f.Message);
                        }

                    }
                    catch (EngineException f)
                    {
                        MessageBox.Show(f.Message);
                    }
                }
                catch (Exception f)
                {

                    MessageBox.Show(f.Message);
                }

                report.Refresh();
                SetCurrentValue(report, "BDName", DBCompany);
                SetCurrentValue(report, "UserCode", UsrCode);
                if (DBServer.Contains("30015") || DBServer.Contains("30013"))
                {
                    #region HANA
                    var dialog = new PrintDialog();

                    var tbloginfo = new TableLogOnInfo();
                    var ci = new ConnectionInfo
                    {
                        DatabaseName = DBCompany,
                        ServerName = "DRIVER={HDBODBC};SERVERNODE={" + DBServer + "};",
                        UserID = DBUserName,
                        Password = DBPassword
                    };
                    tbloginfo.ConnectionInfo = ci;
                    report.Database.Tables[0].ApplyLogOnInfo(tbloginfo);
                    #endregion
                }

                if (impresionAutomatica)
                {
                    if (!printerName.Equals(""))
                    {
                        report.PrintOptions.PrinterName = printerName.Replace("%", " ");
                    }
                    report.PrintToPrinter(1, false, 0, 0);
                    Close();
                    Application.Exit();
                }
                else
                {
                    if (exportar_reporte)
                    {
                        var asda = System.IO.Path.Combine(RepToExport, "56576.doc");
                        report.ExportToDisk(ExportFormatType.WordForWindows, asda);

                        Close();
                        Application.Exit();
                    }
                    else
                    {
                        
                        crystalReportViewer.ReportSource = report;
                        crystalReportViewer.Show();
                    }

                }
            }
            catch (Exception engEx)
            {
                MessageBox.Show("Parámetros de conexión incorrectos." + engEx.Message);
                this.Close();
                Application.Exit();
            }
            finally
            {

            }
            return 0;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                ConfigureCrystalReports();
            }
            catch (Exception engEx)
            {

            }
            finally
            {
                
            }
        }

        private void SetDBLogonForReport(ConnectionInfo connectionInfo, ReportDocument reportDoc)
        {
            try
            {
                Tables tables = reportDoc.Database.Tables;
                foreach (Table table in tables)
                {
                    TableLogOnInfo tableLogOnInfo = table.LogOnInfo;
                    tableLogOnInfo.ConnectionInfo = connectionInfo;
                    table.ApplyLogOnInfo(tableLogOnInfo);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excepción cargando conexiones a las tablas");
            }
        }

        private void crystalReportViewer_HandleException(object source, CrystalDecisions.Windows.Forms.ExceptionEventArgs e)
        {
            if (e.Exception is EngineException)
            {
                EngineException engEx = (EngineException)e.Exception;
                if (engEx.ErrorID == EngineExceptionErrorID.DataSourceError)
                {
                    e.Handled = true;
                    MessageBox.Show("Parámetros de conexión incorrectos. CR003. ");
                }
                else if (engEx.ErrorID == EngineExceptionErrorID.LogOnFailed)
                {
                    e.Handled = true;
                    MessageBox.Show("Parámetros de conexión incorrectos. CR004. ");
                }
                else
                    MessageBox.Show("Error Inesperado. CR005.");
            }
        }

        private void SetCurrentValue(ReportDocument reportDocument, string parameter, string name)
        {
            ParameterFieldDefinitions parameterFieldDefinitions = reportDocument.DataDefinition.ParameterFields;
            ParameterDiscreteValue parameterDiscreteValue;
            ParameterFieldDefinition parameterFieldDefinition;
            ParameterValues currentParameterValues = new ParameterValues();

            parameterDiscreteValue = new ParameterDiscreteValue();
            parameterDiscreteValue.Value = name.ToString();
            currentParameterValues.Add(parameterDiscreteValue);
            parameterFieldDefinition = parameterFieldDefinitions[parameter];
            parameterFieldDefinition.ApplyCurrentValues(currentParameterValues);
           
        }

        private void SetCurrentValuesForParameterField(ReportDocument reportDocument, ArrayList arrayList)
        {
            ParameterFieldDefinitions parameterFieldDefinitions = reportDocument.DataDefinition.ParameterFields;
            ParameterDiscreteValue parameterDiscreteValue;
            ParameterFieldDefinition parameterFieldDefinition;
            ParameterValues currentParameterValues = new ParameterValues();

            int numParam = 1;
            string nomParam;
            foreach (object submittedValue in arrayList)
            {
                //        MessageBox.Show(submittedValue.ToString());
                parameterDiscreteValue = new ParameterDiscreteValue();
                parameterDiscreteValue.Value = submittedValue.ToString();
                currentParameterValues.Add(parameterDiscreteValue);
                if (numParam < 10)
                    nomParam = "Param0" + numParam;
                else
                    nomParam = "Param" + numParam;
                //        MessageBox.Show(nomParam);
                numParam = numParam + 1;
                parameterFieldDefinition = parameterFieldDefinitions[nomParam];
                parameterFieldDefinition.ApplyCurrentValues(currentParameterValues);
            }  
        }
    }
}

