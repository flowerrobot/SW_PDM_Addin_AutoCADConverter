using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IAC_PDM_Addin
{
    //Wrapper class to use SOLIDWORKS PDM Professional as the parent window
   public class WindowHandle : IWin32Window

    {
        private IntPtr mHwnd;

        public WindowHandle(int hWnd)
        {
            mHwnd = new IntPtr(hWnd);
        }
        public IntPtr Handle
        {
            get { return mHwnd; }
        }

    }
}
