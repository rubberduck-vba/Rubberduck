using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.VbeRuntime
{
    internal class VbeNativeApi7 : IVbeNativeApi
    {
        private const string _dllName = "vbe7.dll";

        public string DllName => _dllName;

        [DllImport(_dllName)]
        private static extern int rtcDoEvents();
        public int DoEvents()
        {
            return rtcDoEvents();
        }
    }
}
