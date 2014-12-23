using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System;

namespace Rubberduck
{
    [ComVisible(false)]
    public class PublisherApp : HostApplicationBase<Application>
    {
        public PublisherApp() : base("Publisher") { }

        public override void Run(string target)
        {
            throw new NotImplementedException("Unit Testing not supported in Publisher");
        }

        protected override string GenerateFullyQualifiedName(string projectName, string moduleName, string methodName)
        {
            return string.Concat(moduleName, ".", methodName);
        }
    }
}