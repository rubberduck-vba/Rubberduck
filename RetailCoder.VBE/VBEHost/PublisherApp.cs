using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System;

namespace Rubberduck
{
    public class PublisherApp : HostApplicationBase<Application>
    {
        public PublisherApp() : base("Publisher") { }

        public override void Run(string target)
        {
            //Publisher does not support the Run method
            throw new NotImplementedException("Unit Testing not supported for Publisher");
        }

        protected override string GenerateFullyQualifiedName(string projectName, string moduleName, string methodName)
        {
            return string.Concat(moduleName, ".", methodName);
        }
    }
}