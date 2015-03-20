using System;
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
