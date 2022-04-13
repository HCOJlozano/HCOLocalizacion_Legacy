using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T1.B1.IvaCosto
{
    public class EventInfoClass
    {
        public string ColUID { get; set; }
        public string FormUID { get; set; }
        public string ItemUID { get; set; }
        public int Row { get; set; }
    }




    public class AddDocumentInfoArgs
    {
        public string ObjectType { get; set; }
        public string ObjectKey { get; set; }
        public string FormtTypeEx { get; set; }
        public string FormUID { get; set; }

    }

}
