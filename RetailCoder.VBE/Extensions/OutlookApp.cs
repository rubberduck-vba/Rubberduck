using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Extensions
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