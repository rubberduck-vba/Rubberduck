using System;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.VbeRuntime;

namespace Rubberduck.VBEditor.VBERuntime
{
    internal class Vb6NativeApi : IVbeNativeApi
    {
        // This is the correct dll - why MS named this vba6 instead of vb6 is beyond me.
        // vbe6.dll already taken? Oh, I know, lets confuse the f*ck out of everyone and use vba6.
        private const string _dllName = "vba6.dll";

        public string DllName => _dllName;

        [DllImport(_dllName)]
        private static extern int rtcDoEvents();
        public int DoEvents()
        {
            return rtcDoEvents();
        }
    }
}
