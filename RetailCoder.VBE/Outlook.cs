using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;

namespace Rubberduck
{
    [ComVisible(false)]
    public class OutlookApp : HostApplicationBase<Application>
    {
        public OutlookApp() : base("Outlook") { }

        public override void Run(string target)
        {
        }

        protected override string GenerateFullyQualifiedName(string projectName, string moduleName, string methodName)
        {
            return string.Concat(moduleName, ".", methodName);
        }
    }
}
