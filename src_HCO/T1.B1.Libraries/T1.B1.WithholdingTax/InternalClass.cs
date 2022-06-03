using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T1.B1.WithholdingTax
{
    public class EventInfoClass
    {
        public string ColUID { get; set; }
        public string FormUID { get; set; }
        public string ItemUID { get; set; }
        public int Row { get; set; }
    }

    public class WithholdingTaxDetail
    {
        public string WTCode { get; set; }
        public double Rate { get; set; }
        public string MMCode { get; set; }
        public double MinBase { get; set; }
        public int WTType { get; set; }
        public string MunGroup { get; set; }
        public string Area { get; set; }
        public double NetBase { get; set; }
        public double VatBase { get; set; }
        public bool isMinBaseValid { get { return MinBase <= (WTType == 1 ? VatBase : NetBase); }}
        public bool assigned { get; set; }
        public List<WithholdingTaxConfigMun> Municipios { get; set; }    

    }

   public class WithholdingTaxConfigMun
    {
        public string MunCode { get; set; }
        public string MunName { get; set; }
    }

    public class B1WithHoldingInfoMatrixLine
    {
        public double BaseAmount { get; set; }
        public double WTAmount { get; set; }
    }

    

    public class AddDocumentInfoArgs
    {
        public string ObjectType { get; set; }
        public string ObjectKey { get; set; }
        public string FormtTypeEx { get; set; }
        public string FormUID { get; set; }

    }

    public class InternalRegistryWTData
    {
        public double WTAmount { get; set; }
        public double PercentFromCode { get; set; }
        public double WT { get; set; }
    }
}
