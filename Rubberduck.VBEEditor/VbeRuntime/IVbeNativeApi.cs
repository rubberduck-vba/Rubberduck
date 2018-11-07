using System;

namespace Rubberduck.VBEditor.VbeRuntime
{
    public interface IVbeNativeApi
    {
        string DllName { get; }
        int DoEvents();
        float GetTimer();
        void GetDateVar(out object retval);
        void GetPresentDate(out object retVal);
        double Shell(IntPtr pathname, short windowstyle);
        void GetTimeVar(out object retVal);
        void ChangeDir(IntPtr path);
        void ChangeDrive(IntPtr driveletter);
        void KillFiles(IntPtr pathname);
        void MakeDir(IntPtr path);
        void RemoveDir(IntPtr path);
        void Beep();
    }
}
