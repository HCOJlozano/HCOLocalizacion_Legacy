using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace T1.B1.RelatedParties
{
    class WindowWrapper : IWin32Window
    {
        private IntPtr _hwnd;

        // Property
        public virtual IntPtr Handle
        {
            get { return _hwnd; }
        }

        // Constructor
        public WindowWrapper(IntPtr handle)
        {
            _hwnd = handle;
        }
    }
}
