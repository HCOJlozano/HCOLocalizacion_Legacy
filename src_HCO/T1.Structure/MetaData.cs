using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using log4net;
using Newtonsoft.Json.Linq;
using SAPbobsCOM;
using T1.Queries;
using T1.Structure.Entities;

namespace T1.Structure
{
    public class MetaData
    {
        private Company oCmpny;
        private string PathToJson;
        private string JsonContent;
        private readonly string USERTABLE_NODE = "UserTables";
        private readonly string USERFIELDS_NODE = "UserFields";
        private readonly string UDO_NODE = "UDO";
        private static readonly string LOG_LEVEL = "Debug";
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, LOG_LEVEL);

        public MetaData(Company cmpy, string PathJsonToRead)
        {
            oCmpny = cmpy;
            PathToJson = PathJsonToRead;

            if (!PathToJson.Equals(string.Empty))
                JsonContent = File.ReadAllText(PathToJson);
        }

        public void CreateStructure()
        {
            CreateNewTables();
            CreateNewFields();
            CreateNewUDOs();
        }

        private void CreateNewTables()
        {
            var actualUserTables = GetCurrentValues(Instance.Queries().Get("GetHCO_Tables"));
            var newTables = ReadValuesFromJson<Entities.UserTables>(USERTABLE_NODE);
            var realNewTables = from t in newTables
                                where !(from a in actualUserTables select a.TableCode).Contains(t.TableCode)
                                select t;

            string sErrMsg = string.Empty;
            UserTablesMD oUserTablesMD = null;
            GC.Collect();

            try
            {
                foreach (Entities.UserTables table in realNewTables)
                {
                    oUserTablesMD = (UserTablesMD)oCmpny.GetBusinessObject(BoObjectTypes.oUserTables);
                    oUserTablesMD.TableName = table.TableCode;
                    oUserTablesMD.TableDescription = table.Description;
                    oUserTablesMD.TableType = (BoUTBTableType)table.Type;
                    
                    if (oUserTablesMD.Add() != 0)
                        sErrMsg += "Error creando campo: " + table.TableCode + " - " + oCmpny.GetLastErrorDescription() + ".\n";
                }

                GC.Collect();
            }
            catch (Exception ex)
            {
                _Logger.Error("COM Error", ex);
            }
            finally
            {
                if (oUserTablesMD != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                oUserTablesMD = null;
                GC.Collect();
            }

            if (!sErrMsg.Equals(string.Empty))
                _Logger.Info("OK");
        }

        private void CreateNewFields()
        {
            var newUserFields = ReadValuesFromJson<Entities.UserFields>(USERFIELDS_NODE);
            var actualUserFields = GetActualUserFields(Instance.Queries().Get("GetUserFields"));

            var tables = from u in newUserFields
                         join a in actualUserFields on new { u.TableCode, u.FieldCode } equals new { a.TableCode, a.FieldCode } into f
                         from uf in f.DefaultIfEmpty()
                         select new { u.TableCode, u.FieldCode, ETableCode = (uf == null ? string.Empty : uf.TableCode), EUserField = (uf == null ? string.Empty : uf.FieldCode) };

            var newFields = from u in tables
                            where u.ETableCode == string.Empty || u.EUserField == string.Empty
                            select new { u.TableCode, u.FieldCode };

            var realNewUserFields = from t in newUserFields
                                    join a in newFields on new { t.TableCode, t.FieldCode } equals new { a.TableCode, a.FieldCode }
                                    select t;

            string sErrMsg = string.Empty;
            UserFieldsMD oUserFieldsMD = null;
            GC.Collect();

            try
            {
                foreach (Entities.UserFields field in realNewUserFields)
                {
                    oUserFieldsMD = (UserFieldsMD)oCmpny.GetBusinessObject(BoObjectTypes.oUserFields);
                    oUserFieldsMD.TableName = field.TableCode;
                    oUserFieldsMD.Name = field.FieldCode;
                    oUserFieldsMD.Description = field.FieldName;
                    oUserFieldsMD.Type = (BoFieldTypes)field.Type;
                    
                    oUserFieldsMD.SubType = (BoFldSubTypes)field.SubType;

                    switch (field.LinkType)
                    {
                        case 0:
                            oUserFieldsMD.LinkedTable = field.LinkCode;
                            break;
                        case 1:
                            oUserFieldsMD.LinkedSystemObject = (UDFLinkedSystemObjectTypesEnum)(Int32.Parse(field.LinkCode));
                            break;
                        case 2:
                            oUserFieldsMD.LinkedUDO = field.LinkCode;
                            break;
                    }

                    if ((oUserFieldsMD.Type != BoFieldTypes.db_Date)) oUserFieldsMD.EditSize = field.Length;
                    if (field.ValidValues.Count > 0)
                    {
                        foreach (Entities.ValidValues val in field.ValidValues)
                        {
                            oUserFieldsMD.ValidValues.SetCurrentLine(oUserFieldsMD.ValidValues.Count - 1);
                            oUserFieldsMD.ValidValues.Value = val.ValidValue;
                            oUserFieldsMD.ValidValues.Description = val.Description;
                            oUserFieldsMD.ValidValues.Add();
                        }
                    }
                    if (!string.IsNullOrEmpty(field.DefaultValue))
                        oUserFieldsMD.DefaultValue = field.DefaultValue;

                    if (oUserFieldsMD.Add() != 0)
                        sErrMsg += "Error creando campo: " + field.FieldCode + " en tabla " + field.TableCode + oCmpny.GetLastErrorDescription() + ".\n";
                }
            }
            catch (Exception ex)
            {
                _Logger.Error("COM Error", ex);
            }
            finally
            {
                if (oUserFieldsMD != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();
            }

            if (!sErrMsg.Equals(string.Empty))
                _Logger.Info("OK");
        }

        private void CreateNewUDOs()
        {
            var actualUDOs = GetActualUDO(Instance.Queries().Get("GetUDOs"));
            var newUDOs = ReadValuesFromJson<UDO>(UDO_NODE);

            var realNewUDO = from t in newUDOs
                             where !(from a in actualUDOs select a.Code).Contains(t.Code)
                             select t;

            string sErrMsg = string.Empty;

            try
            {
                foreach (UDO oUDO in realNewUDO)
                {
                    var MyUDO = (UserObjectsMD)oCmpny.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
                    MyUDO.Code = oUDO.Code;
                    MyUDO.Name = oUDO.Description;
                    MyUDO.ObjectType = (BoUDOObjType)oUDO.Type;
                    MyUDO.TableName = oUDO.TableCode;

                    // Set Services
                    MyUDO.CanCancel = BoYesNoEnum.tYES;
                    MyUDO.CanClose = BoYesNoEnum.tYES;
                    MyUDO.CanDelete = oUDO.CanDelete == "Y" ? BoYesNoEnum.tYES : BoYesNoEnum.tNO;

                    if (oUDO.DefaultForm == 1)
                    {
                        MyUDO.MenuItem = BoYesNoEnum.tYES;
                        MyUDO.CanCreateDefaultForm = BoYesNoEnum.tYES;
                        MyUDO.EnableEnhancedForm = BoYesNoEnum.tNO;
                        MyUDO.MenuCaption = oUDO.MenuCaption;
                        MyUDO.MenuUID = oUDO.MenuID;
                        MyUDO.FatherMenuID = oUDO.FatherMenuID;
                        MyUDO.Position = oUDO.Position;
                        foreach (FormColumns col in oUDO.FormColumns)
                        {
                            MyUDO.FormColumns.Add();
                            MyUDO.FormColumns.FormColumnAlias = col.Column;
                            MyUDO.FormColumns.FormColumnDescription = col.Description;
                            MyUDO.FormColumns.Editable = BoYesNoEnum.tYES;
                        }
                    }
                    else
                        MyUDO.CanCreateDefaultForm = BoYesNoEnum.tNO;

                   
                    MyUDO.CanFind = BoYesNoEnum.tYES;
                    MyUDO.CanLog = BoYesNoEnum.tYES;
                    MyUDO.CanYearTransfer = BoYesNoEnum.tNO;
                    MyUDO.ManageSeries = BoYesNoEnum.tNO;
                    MyUDO.RebuildEnhancedForm = BoYesNoEnum.tNO;

                    if (MyUDO.ObjectType.Equals(BoUDOObjType.boud_MasterData))
                    {
                        MyUDO.FindColumns.ColumnAlias = "Code";
                        MyUDO.FindColumns.ColumnDescription = "Code";
                        MyUDO.FindColumns.Add();
                        MyUDO.FindColumns.ColumnAlias = "Name";
                        MyUDO.FindColumns.ColumnDescription = "Descripcion";
                        MyUDO.FindColumns.Add();
                    }

                    if (MyUDO.ObjectType.Equals(BoUDOObjType.boud_Document))
                    {
                        MyUDO.FindColumns.ColumnAlias = "DocEntry";
                        MyUDO.FindColumns.ColumnDescription = "DocEntry";
                        MyUDO.FindColumns.Add();
                    }

                    if (oUDO.ChildTables.Count > 0)
                    {
                        foreach (ChildTables Table in oUDO.ChildTables)
                        {
                            MyUDO.ChildTables.TableName = Table.ChildTableCode;
                            MyUDO.ChildTables.Add();
                        }
                    }

                    if ((MyUDO.Add() != 0))
                        sErrMsg += "Error creando UDO: " + oUDO.Code + " - " + oCmpny.GetLastErrorDescription() + ".\n";

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(MyUDO);
                    MyUDO = null;
                }
            }
            catch (Exception ex)
            {
                _Logger.Error("COM Error", ex);
            }

            if (!sErrMsg.Equals(string.Empty))
                _Logger.Info("OK");
        }

        private List<T> ReadValuesFromJson<T>(string nodeToRead)
        {
            JObject o = JObject.Parse(JsonContent);
            JArray a = (JArray)o[nodeToRead];
            return a.ToObject<List<T>>();
        }

        public List<Entities.UserTables> GetCurrentValues(string query)
        {
            var userTables = new List<Entities.UserTables>();
            var record = (Recordset)oCmpny.GetBusinessObject(BoObjectTypes.BoRecordset);
            record.DoQuery(query);

            while (!record.EoF)
            {
                userTables.Add(new Entities.UserTables { TableCode = record.Fields.Item("TableName").Value.ToString() });
                record.MoveNext();
            }

            return userTables;
        }

        public List<Entities.UserFields> GetActualUserFields(string query)
        {
            var userFields = new List<Entities.UserFields>();
            var record = (Recordset)oCmpny.GetBusinessObject(BoObjectTypes.BoRecordset);
            record.DoQuery(query);

            while (!record.EoF)
            {
                userFields.Add(new Entities.UserFields { TableCode = record.Fields.Item("TableID").Value.ToString(), FieldCode = record.Fields.Item("AliasID").Value.ToString() });
                record.MoveNext();
            }

            return userFields;
        }

        public List<UDO> GetActualUDO(string query)
        {
            var userFields = new List<UDO>();
            var record = (Recordset)oCmpny.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                record.DoQuery(query);

                while (!record.EoF)
                {
                    userFields.Add(new UDO { Code = record.Fields.Item("Code").Value.ToString() });
                    record.MoveNext();
                }

                return userFields;
            }
            catch { return userFields; }
            finally
            {
                if (record != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(record);

                record = null;
                GC.Collect();
            }
        }

    }
}
