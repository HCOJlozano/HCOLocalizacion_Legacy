using System;
using System.Windows.Forms;

namespace T1.CrystalReport
{
    static class Program
  {
    [STAThread]
    static void Main(string[] args)
    {
      if (args.Length == 0 ) return;


      Application.EnableVisualStyles();
      Application.SetCompatibleTextRenderingDefault(false);
      Application.Run(new Form1(args[0], args[1], args[2], args[3], args[4], args[5], args[6]));
    }
  }
}