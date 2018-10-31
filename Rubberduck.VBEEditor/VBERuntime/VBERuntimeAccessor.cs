using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.VBERuntime
{
    public class VBERuntimeAccessor : IVBERuntime
    {
        private static DllVersion _version;
        private readonly IVBERuntime _runtime;
        
        static VBERuntimeAccessor()
        {
            _version = DllVersion.Unknown;
        }
        
        public VBERuntimeAccessor(IVBE vbe)
        {
            if (_version == DllVersion.Unknown)
            {
                try
                {
                    _version = VBEDllVersion.GetCurrentVersion(vbe);
                }
                catch
                {
                    _version = DllVersion.Unknown;
                }
            }
            _runtime = InitializeRuntime();
        }

        private static IVBERuntime InitializeRuntime()
        {
            switch (_version)
            {
                case DllVersion.Vbe7:
                    return new VBERuntime7();
                case DllVersion.Vbe6:
                    return new VBERuntime6();
                default:
                    return DetermineVersion();
            }
        }

        private static IVBERuntime DetermineVersion()
        {
            IVBERuntime runtime;
            try
            {
                runtime = new VBERuntime7();
                runtime.GetTimer();
                _version = DllVersion.Vbe7;
            }
            catch
            {
                try
                {
                    runtime = new VBERuntime6();
                    runtime.GetTimer();
                    _version = DllVersion.Vbe6;
                }
                catch
                {
                    // we shouldn't be here.... Rubberduck is a VBA add-in, so how the heck could it have loaded without a VBE dll?!?
                    throw new InvalidOperationException("Cannot execute DoEvents; the VBE dll could not be located.");
                }
            }

            return _version != DllVersion.Unknown ? runtime : null;
        }

        public string DllName => _runtime.DllName;

        public float GetTimer()
        {
            return _runtime.GetTimer();
        }

        public void GetDateVar(out object retval)
        {
            _runtime.GetDateVar(out retval);
        }

        public void GetPresentDate(out object retVal)
        {
            _runtime.GetPresentDate(out retVal);
        }

        public double Shell(IntPtr pathname, short windowstyle)
        {
            return _runtime.Shell(pathname, windowstyle);
        }

        public void GetTimeVar(out object retVal)
        {
            _runtime.GetTimeVar(out retVal);
        }

        public void ChangeDir(IntPtr path)
        {
            _runtime.ChangeDir(path);
        }

        public void ChangeDrive(IntPtr driveletter)
        {
            _runtime.ChangeDrive(driveletter);
        }

        public void KillFiles(IntPtr pathname)
        {
            _runtime.KillFiles(pathname);
        }

        public void MakeDir(IntPtr path)
        {
            _runtime.MakeDir(path);
        }

        public void RemoveDir(IntPtr path)
        {
            _runtime.RemoveDir(path);
        }

        public int DoEvents()
        {
            return _runtime.DoEvents();
        }

        public void Beep()
        {
            _runtime.Beep();
        }
    }
}
