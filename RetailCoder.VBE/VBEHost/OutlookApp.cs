using System;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;

namespace Rubberduck.VBEHost
{
    public class OutlookApp : HostApplicationBase<Application>
    {
        public OutlookApp() : base("Outlook") { }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            //Outlook does not support the run method.
            throw new NotImplementedException("Unit Testing not supported for Outlook");
        }

        protected virtual string GenerateMethodCall(string projectName, string moduleName, string methodName)
        {
            return string.Concat(moduleName, ".", methodName);
        }
    }
}