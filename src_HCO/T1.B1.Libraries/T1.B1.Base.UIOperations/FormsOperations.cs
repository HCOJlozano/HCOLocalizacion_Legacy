using log4net;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;


namespace T1.B1.Base.UIOperations
{
    public class FormsOperations
    {
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);

        private FormsOperations()
        {

        }

        public static void AddChooseFromList(Form oForm, string objType, string uniqueID, bool MultiSelc)
        {
            try
            {
                ChooseFromListCollection oCFLs = null;
                Conditions oCons = null;
                Condition oCon = null;

                oCFLs = oForm.ChooseFromLists;

                SAPbouiCOM.ChooseFromList oCFL = null;
                ChooseFromListCreationParams oCFLCreationParams = null;
                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(B1.MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_ChooseFromListCreationParams)));

                oCFLCreationParams.MultiSelection = MultiSelc;
                oCFLCreationParams.ObjectType = objType;
                oCFLCreationParams.UniqueID = uniqueID;

                oCFL = oCFLs.Add(oCFLCreationParams);
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void SetChooseFromList(Form oForm, string CFL_ID, string alias, SAPbouiCOM.BoConditionOperation operation, string condVal)
        {
            try
            {
                var oCFL = oForm.ChooseFromLists.Item(CFL_ID);
                var oCons = oCFL.GetConditions();
                var oCon = oCons.Add();
                oCon.Alias = alias;
                oCon.Operation = operation;
                oCon.CondVal = condVal;
                oCFL.SetConditions(oCons);
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void AddRightClickMenu(string uniqueID, string title, int position)
        {
            if (MainObject.Instance.B1Application.Menus.Exists(uniqueID)) return;

            SAPbouiCOM.MenuItem oMenuItem = null;
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuCreationParams oCreationPackage = null;

            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            try
            {
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = uniqueID;
                oCreationPackage.String = title;
                oCreationPackage.Enabled = true;
                oCreationPackage.Position = position;
                oMenuItem = MainObject.Instance.B1Application.Menus.Item("1280");
                oMenus = oMenuItem.SubMenus;
                oMenus.AddEx(oCreationPackage);
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }

        }

        public static void DeleteRightClickMenu(string uniqueID)
        {
            try
            {
                if (MainObject.Instance.B1Application.Menus.Exists(uniqueID))
                    MainObject.Instance.B1Application.Menus.RemoveEx(uniqueID);
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void UpdateMatrixRowNumbers(string ItemID, Form oForm)
        {
            var matrix = (Matrix)oForm.Items.Item(ItemID).Specific;
            for (int i = 1; i <= matrix.RowCount; i++)
                ((EditText)matrix.GetCellSpecific("#", i)).Value = i.ToString();
        }

        public static System.Data.DataTable SapDataTableToDotNetDataTable(string XMLDatatable)
        {
            var DT = new System.Data.DataTable();
            var XDoc = System.Xml.Linq.XDocument.Parse(XMLDatatable);
            var Columns = XDoc.Element("DataTable").Element("Columns").Elements("Column");

            foreach (var Column in Columns)
            {
                DT.Columns.Add(Column.Attribute("Uid").Value, ((Column.Attribute("Type").Value.ToString().Equals("5") || Column.Attribute("Type").Value.ToString().Equals("8")) ? typeof(System.Double) : typeof(System.String)));
            }

            var Rows = XDoc.Element("DataTable").Element("Rows").Elements("Row");

            //var Names = new List<string>();
            foreach (var Row in Rows)
            {
                var DTRow = DT.NewRow();

                var Cells = Row.Element("Cells").Elements("Cell");

                foreach (var Cell in Cells)
                {
                    var ColName = Cell.Element("ColumnUid").Value;
                    var ColValue = Cell.Element("Value").Value;
                    if (DT.Columns[ColName].DataType.Name.Equals("Double")) DTRow[ColName] = double.Parse(ColValue, System.Globalization.CultureInfo.InvariantCulture);
                    else DTRow[ColName] = ColValue;
                }

                DT.Rows.Add(DTRow);
            }

            return DT;
        }

        public static System.Data.DataTable SapDBDataSourceToDotNetDataTable(SAPbouiCOM.DBDataSource sap_table)
        {
            var DT = new System.Data.DataTable();

            for (int i = 0; i < sap_table.Fields.Count; i++)
            {
                DT.Columns.Add(sap_table.Fields.Item(i).Name, (sap_table.Fields.Item(i).Type == BoFieldsType.ft_Float ? typeof(System.Double) : typeof(System.String)));
            }

            var XDoc = System.Xml.Linq.XDocument.Parse(sap_table.GetAsXML()); //System.Xml.Linq.XDocument.Parse(XMLDatatable);

            var Rows = XDoc.Element("dbDataSources").Element("rows").Elements("row");

            //var Names = new List<string>();
            foreach (var Row in Rows)
            {
                var DTRow = DT.NewRow();

                var Cells = Row.Element("cells").Elements("cell");

                foreach (var Cell in Cells)
                {
                    var ColName = Cell.Element("uid").Value;
                    var ColValue = Cell.Element("value").Value;
                    if (DT.Columns[ColName].DataType.Name.Equals("Double")) DTRow[ColName] = double.Parse(ColValue, System.Globalization.CultureInfo.InvariantCulture);
                    else DTRow[ColName] = ColValue;
                }

                DT.Rows.Add(DTRow);
            }

            return DT;
        }

        public static object[] ListChoiceListener(ItemEvent pVal, string columna)
        {
            try
            {
                var oCflEvento = (IChooseFromListEvent)pVal;
                if (oCflEvento.BeforeAction) return new object[] { };

                var oDataTable = oCflEvento.SelectedObjects;
                if (oDataTable == null)
                    return new object[] { "" };

                var listRows = new object[oDataTable.Rows.Count];

                if (!(!pVal.BeforeAction & oCflEvento.SelectedObjects != null)) return new object[] { };

                if (oDataTable.Rows.Count > 1)
                    for (int i = 0; i < oDataTable.Rows.Count; i++)
                        listRows[i] = oDataTable.GetValue(columna, i);
                else
                    return new object[] { oDataTable.GetValue(columna, 0) };

                return listRows;
            }
            catch (Exception) { }
            return new object[] { };
        }

    }
}
