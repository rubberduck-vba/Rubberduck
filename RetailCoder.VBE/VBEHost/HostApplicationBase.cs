using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Rubberduck
{
    public abstract class HostApplicationBase<TApplication> : IHostApplication
    {
        protected readonly TApplication _application;
        protected HostApplicationBase(string applicationName)
        {
            _application = (TApplication)Marshal.GetActiveObject(applicationName + ".Application");
        }

        ~HostApplicationBase()
        {
            Marshal.ReleaseComObject(_application);
        }

        public abstract void Run(string target);
        protected abstract string GenerateFullyQualifiedName(string projectName, string moduleName, string methodName);

        public TimeSpan TimedMethodCall(string projectName, string moduleName, string methodName)
        {
            var stopwatch = Stopwatch.StartNew();

            Run(GenerateFullyQualifiedName(projectName, moduleName, methodName));

            stopwatch.Stop();
            return stopwatch.Elapsed;
        }
    }
}