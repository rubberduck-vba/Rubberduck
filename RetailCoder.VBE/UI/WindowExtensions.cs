using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI
{
    public static class WindowExtensions
    {
        public static IntPtr Handle(this Window window)
        {
           return (IntPtr)window.HWnd;
        }
    }
}
