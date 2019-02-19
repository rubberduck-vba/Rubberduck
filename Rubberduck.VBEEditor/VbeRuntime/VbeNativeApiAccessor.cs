using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using NLog;
using Rubberduck.VBEditor.VBERuntime;

namespace Rubberduck.VBEditor.VbeRuntime
{
    public class VbeNativeApiAccessor : IVbeNativeApi
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private static readonly DllVersion Version;
        private IVbeNativeApi _runtime;
        
        static VbeNativeApiAccessor()
        {
            Version = DllVersion.Unknown;
        }
        
        private static readonly List<(string Name, Type ApiType, DllVersion Version)> VbeApis =new List<(string Dll, Type ApiType, DllVersion Version)>
        {
            ( "vbe6.dll", typeof(VbeNativeApi6), DllVersion.Vbe6),
            ( "vbe7.dll", typeof(VbeNativeApi7), DllVersion.Vbe7),
            ( "vba6.dll", typeof(Vb6NativeApi), DllVersion.Vb98)
        };

        public VbeNativeApiAccessor()
        {
            DetermineVersion();
        }

        private IVbeNativeApi DetermineVersion()
        {
            try
            {
                var modules = Process.GetCurrentProcess().Modules.OfType<ProcessModule>()
                    .Select(module => module.ModuleName.ToLowerInvariant())
                    .ToList();

                var api = VbeApis.FirstOrDefault(dll => modules.Contains(dll.Name));

                if (api.ApiType is null)
                {
                    // Yay! VBE8 must have been released! Duck out.
                    throw new InvalidOperationException("Cannot execute library function; the VBE dll could not be located.");
                }

                _runtime = (IVbeNativeApi)Activator.CreateInstance(api.ApiType);
            }
            catch (Exception ex)
            {
                Logger.Warn(ex, "Exception during location of the VBE dll version. Resolution deferred.");
            }

            return null;
        }

        public string DllName => _runtime?.DllName ?? DetermineVersion()?.DllName;

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
