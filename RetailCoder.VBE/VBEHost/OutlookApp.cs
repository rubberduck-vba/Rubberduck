using System;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEHost
{
    public class OutlookApp : HostApplicationBase<Application>
    {
        public OutlookApp() : base("Outlook") { }

        public override void Run(string target)
        {
            //Outlook does not support the run method.
            throw new NotImplementedException("Unit Testing not supported for Publisher");
        }

        protected override string GenerateFullyQualifiedName(string projectName, string moduleName, string methodName)
        {
            return string.Concat(moduleName, ".", methodName);
        }
    }
}