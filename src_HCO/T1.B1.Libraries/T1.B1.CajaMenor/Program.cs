using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;

namespace T1.B1.CajaMenor
{
    class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                Application oApp = null;

                oApp.Run();
            }
            catch (Exception ex)
            {
                
            }
        }
    }
}
