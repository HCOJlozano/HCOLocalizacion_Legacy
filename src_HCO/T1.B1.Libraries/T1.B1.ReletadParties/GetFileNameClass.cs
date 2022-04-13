using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace T1.B1.RelatedParties
{
    class GetFileNameClass
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        OpenFileDialog _oFileDialog;

        // Properties
        public string Path
        {
            get { return _oFileDialog.FileName; }
            set { _oFileDialog.FileName = value; }
        }

        // Constructor
        public GetFileNameClass()
        {
            _oFileDialog = new OpenFileDialog();
            _oFileDialog.Multiselect = false;
            _oFileDialog.Filter = "Archivos TXT(*.txt)|*.txt";
        }

        // Methods

        public void GetFileName()
        {
            IntPtr ptr = GetForegroundWindow();
            WindowWrapper oWindow = new WindowWrapper(ptr);
            if (_oFileDialog.ShowDialog(oWindow) != DialogResult.OK)
            {
                _oFileDialog.FileName = string.Empty;
            }
            oWindow = null;
        } // End of GetFileName
    }
}
