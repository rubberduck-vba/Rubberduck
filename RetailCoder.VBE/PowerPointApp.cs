using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Rubberduck
{
    [ComVisible(false)]
    public class PowerPointApp : HostApplicationBase<Application>
    {
        public PowerPointApp() : base("PowerPoint") { }

        public override void Run(string target)
        {
            object[] o = { }; //powerpoint requires a paramarray, so we pass it an empty array.
            base._application.Run(target, o);
        }

        protected override string GenerateFullyQualifiedName(string projectName, string moduleName, string methodName)
        {
            return string.Concat(moduleName, ".", methodName);
        }
    }
}