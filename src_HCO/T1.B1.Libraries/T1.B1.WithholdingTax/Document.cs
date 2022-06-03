using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T1.B1.WithholdingTax
{
    public class Document
    {
        string cardCode;
        string Withholding;
        int lines;
        string formid;



    }

    public class DocumentLines
    {
        string itemCode;
        double LineTotal;
        double vatSum;
        string taxableSum;
    }

}
