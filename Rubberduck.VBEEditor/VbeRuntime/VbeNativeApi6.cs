using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.VbeRuntime
{
    internal class VbeNativeApi6 : IVbeNativeApi
    {
        private const string _dllName = "vbe6.dll";

        public string DllName => _dllName;

        [DllImport(_dllName)]
        private static extern int rtcDoEvents();
        public int DoEvents()
        {
            return rtcDoEvents();
        }
    }
}
