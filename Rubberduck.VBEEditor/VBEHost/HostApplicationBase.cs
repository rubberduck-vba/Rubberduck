using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.VBEHost
{
    [ComVisible(false)]
    public abstract class HostApplicationBase<TApplication> : IHostApplication
        where TApplication : class
    {
        private readonly string _applicationName;
        protected readonly TApplication Application;
        protected HostApplicationBase(string applicationName)
        {
            _applicationName = applicationName;

            try
            {
                Application = (TApplication)Marshal.GetActiveObject(applicationName + ".Application");
            }
            catch (COMException)
            {
                Application = null; // unit tests don't need it anyway.
            }
        }

        ~HostApplicationBase()
        {
            if (Application != null)
            {
                Marshal.ReleaseComObject(Application);
            }
        }

        public string ApplicationName
        {
            get { return _applicationName; }
        }

        public abstract void Run(QualifiedMemberName qualifiedMemberName);

        public abstract void Save();

        public TimeSpan TimedMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            var stopwatch = Stopwatch.StartNew();

            Run(qualifiedMemberName);

            stopwatch.Stop();
            return stopwatch.Elapsed;
        }

        public void Dispose()
        {
            if (Application != null)
            {
                Marshal.ReleaseComObject(Application);
            }
        }
    }
}