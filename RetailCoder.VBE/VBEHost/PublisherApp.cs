using System;
using Microsoft.Office.Interop.Excel;

namespace Rubberduck.VBEHost
{
    public class PublisherApp : HostApplicationBase<Application>
    {
        public PublisherApp() : base("Publisher") { }

        public override void Run(string projectName, string moduleName, string methodName)
        {
            //Publisher does not support the Run method
            throw new NotImplementedException("Unit Testing not supported for Publisher");
        }

        protected override string GenerateMethodCall(string projectName, string moduleName, string methodName)
        {
            return string.Concat(moduleName, ".", methodName);
        }
    }
}