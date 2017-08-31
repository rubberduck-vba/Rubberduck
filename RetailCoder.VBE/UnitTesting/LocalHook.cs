using System;

namespace EasyHook
{
    internal class LocalHook : IDisposable
    {
        public HookAccessControl ThreadACL => new HookAccessControl();

        public void Dispose() { }

        public static LocalHook Create(IntPtr inTargetProc, Delegate inNewProc, object inCallback)
        {
            return new LocalHook();
        }

        public static IntPtr GetProcAddress(string inModule, string inSymbolName)
        {
            return new IntPtr(0);
        }

        public class HookAccessControl
        {
            public void SetInclusiveACL(int[] InACL) { }
        }
    }
}