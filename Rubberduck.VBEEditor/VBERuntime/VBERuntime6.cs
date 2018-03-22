using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.VBERuntime
{
    internal class VBERuntime6 : IVBERuntime
    {
        private const string DllName = "vbe6.dll";

        [DllImport(DllName)]
        private static extern int rtcDoEvents();
        public int DoEvents()
        {
            return rtcDoEvents();
        }

        [DllImport(DllName)]
        private static extern float rtcGetTimer();
        public float Timer()
        {
            return rtcGetTimer();
        }
    }
}
