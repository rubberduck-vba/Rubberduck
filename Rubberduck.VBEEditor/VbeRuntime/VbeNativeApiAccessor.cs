using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.VBERuntime;

namespace Rubberduck.VBEditor.VbeRuntime
{
    public class VbeNativeApiAccessor : IVbeNativeApi
    {
        private static DllVersion _version;
        private readonly IVbeNativeApi _runtime;
        
        static VbeNativeApiAccessor()
        {
            _version = DllVersion.Unknown;
        }
        
        public VbeNativeApiAccessor(IVBE vbe)
        {
            if (_version == DllVersion.Unknown)
            {
                try
                {
                    _version = VbeDllVersion.GetCurrentVersion(vbe);
                }
                catch
                {
                    _version = DllVersion.Unknown;
                }
            }
            _runtime = InitializeRuntime();
        }

        private static IVbeNativeApi InitializeRuntime()
        {
            switch (_version)
            {
                case DllVersion.Vbe7:
                    return new VbeNativeApi7();
                case DllVersion.Vbe6:
                    return new VbeNativeApi6();
                case DllVersion.Vb98:
                    return new Vb6NativeApi();
                default:
                    return DetermineVersion();
            }
        }

        private static IVbeNativeApi DetermineVersion()
        {
            foreach (var type in new[] {typeof(VbeNativeApi7), typeof(VbeNativeApi6), typeof(Vb6NativeApi)})
            {
                try
                {
                    var runtime = (IVbeNativeApi)Activator.CreateInstance(type);
                    runtime.GetTimer();
                    return runtime;
                }
                catch
                {
                    // ignored
                }
            }
            // we shouldn't be here.... Rubberduck is a VBE add-in, so how the heck could it have loaded without a runtime dll?!?
            throw new InvalidOperationException("Cannot execute library function; the VBE dll could not be located.");
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
