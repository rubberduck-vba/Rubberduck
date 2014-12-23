using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Rubberduck
{
    [ComVisible(false)]
    public class PublisherApp : HostApplicationBase<Application>
    {
        public PublisherApp() : base("Publisher") { }

        public override void Run(string target)
        {
        }

        protected override string GenerateFullyQualifiedName(string projectName, string moduleName, string methodName)
        {
            return string.Concat(moduleName, ".", methodName);
        }
    }
}