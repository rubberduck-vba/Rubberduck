using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.VBEHost
{
    public abstract class HostApplicationBase<TApplication> : IHostApplication
        where TApplication : class
    {
        protected readonly TApplication Application;
        protected HostApplicationBase(string applicationName)
        {
            Application = (TApplication)Marshal.GetActiveObject(applicationName + ".Application");
        }

        ~HostApplicationBase()
        {
            if (Application != null)
            {
                Marshal.ReleaseComObject(Application);
            }
        }

        public abstract void Run(QualifiedMemberName qualifiedMemberName);

        public TimeSpan TimedMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            var stopwatch = Stopwatch.StartNew();

            Run(qualifiedMemberName);

            stopwatch.Stop();
            return stopwatch.Elapsed;
        }
    }
}