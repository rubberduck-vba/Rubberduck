using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;

namespace Rubberduck
{
    [ComVisible(false)]
    public class PowerPointApp : HostApplicationBase<Application>
    {
        public PowerPointApp() : base("PowerPoint") { }

        public override void Run(string target)
        {
            object[] o = { };
            base._application.Run(target, o);
        }

        protected override string GenerateFullyQualifiedName(string projectName, string moduleName, string methodName)
        {
            return string.Concat(moduleName, ".", methodName);
        }
    }
}
