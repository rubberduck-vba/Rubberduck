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
            object[] paramArray = { }; //powerpoint requires a paramarray, so we pass it an empty array.
            base._application.Run(target, paramArray);
        }

        protected override string GenerateFullyQualifiedName(string projectName, string moduleName, string methodName)
        {
            /* Note: Powerpoint supports a `FileName.ppt!module.method` syntax, 
             * but that would require significant changes to the Unit Testing Framework.
             * http://msdn.microsoft.com/en-us/library/office/ff744221(v=office.15).aspx
             */

            return string.Concat(moduleName, ".", methodName);
        }
    }
}