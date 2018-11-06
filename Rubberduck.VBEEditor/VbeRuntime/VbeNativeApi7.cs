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

        [DllImport(_dllName, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.R4)]
        private static extern float rtcGetTimer();
        public float GetTimer()
        {
            return rtcGetTimer();
        }

        [DllImport(_dllName, SetLastError = true)]
        private static extern void rtcGetDateVar(out object retVal);
        public void GetDateVar(out object retval)
        {
            rtcGetDateVar(out retval);
        }

        [DllImport(_dllName, SetLastError = true)]
        private static extern void rtcGetPresentDate(out object retVal);
        public void GetPresentDate(out object retVal)
        {
            rtcGetPresentDate(out retVal);
        }

        [DllImport(_dllName, SetLastError = true)]
        private static extern double rtcShell(IntPtr pathname, short windowstyle);
        public double Shell(IntPtr pathname, short windowstyle)
        {
            return rtcShell(pathname, windowstyle);
        }

        [DllImport(_dllName, SetLastError = true)]
        private static extern void rtcGetTimeVar(out object retVal);
        public void GetTimeVar(out object retVal)
        {
            rtcGetTimeVar(out retVal);
        }

        [DllImport(_dllName, SetLastError = true)]
        private static extern void rtcChangeDir(IntPtr path);
        public void ChangeDir(IntPtr path)
        {
            rtcChangeDir(path);
        }

        [DllImport(_dllName, SetLastError = true)]
        private static extern void rtcChangeDrive(IntPtr driveletter);
        public void ChangeDrive(IntPtr driveletter)
        {
            rtcChangeDrive(driveletter);
        }

        [DllImport(_dllName, SetLastError = true)]
        private static extern void rtcKillFiles(IntPtr pathname);
        public void KillFiles(IntPtr pathname)
        {
            rtcKillFiles(pathname);
        }

        [DllImport(_dllName, SetLastError = true)]
        private static extern void rtcMakeDir(IntPtr path);
        public void MakeDir(IntPtr path)
        {
            rtcMakeDir(path);
        }

        [DllImport(_dllName, SetLastError = true)]
        private static extern void rtcRemoveDir(IntPtr path);
        public void RemoveDir(IntPtr path)
        {
            rtcRemoveDir(path);
        }

        [DllImport(_dllName, SetLastError = true)]
        private static extern void rtcBeep();
        public void Beep()
        {
            rtcBeep();
        }
    }
}
