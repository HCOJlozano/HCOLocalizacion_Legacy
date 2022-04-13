using System.Collections.Generic;

namespace T1.Structure.Entities
{
    public class UserTables
    {
        public string TableCode = string.Empty;
        public string Description = string.Empty;
        public int Type = 0;
    }

    public class UserFields
    {
        public string TableCode = string.Empty;
        public string FieldCode = string.Empty;
        public string FieldName = string.Empty;
        public int Type = 0;
        public int SubType = 0;
        public int Length = 0;
        public int LinkType = 0;
        public string LinkCode = string.Empty;
        public string DefaultValue = string.Empty;
        public List<ValidValues> ValidValues = new List<ValidValues>();
    }

    public class ValidValues
    {
        public string ValidValue = string.Empty;
        public string Description = string.Empty;
    }

    public class UDO
    {
        public string Code = string.Empty;
        public string Description = string.Empty;
        public string TableCode = string.Empty;
        public int Type = 0;
        public int DefaultForm = 0;
        public int Position = 0;
        public string MenuID = string.Empty;
        public string MenuCaption = string.Empty;
        public int FatherMenuID = 0;
        public string CanDelete = string.Empty;
        public List<ChildTables> ChildTables = new List<Entities.ChildTables>();
        public List<FormColumns> FormColumns = new List<Entities.FormColumns>();
    }

    public class ChildTables
    {
        public string ChildTableCode = string.Empty;
        public List<FormColumns> FormColumns = new List<Entities.FormColumns>();
    }

    public class FormColumns
    {
        public string Column = string.Empty;
        public string Description = string.Empty;
    }
}
