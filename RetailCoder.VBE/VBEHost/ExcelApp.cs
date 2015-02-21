using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Rubberduck
{
    public class ExcelApp : HostApplicationBase<Application>
    {
        public ExcelApp() : base("Excel") { }

        public override void Run(string target)
        {
            base._application.Run(target);
        }

        protected override string GenerateFullyQualifiedName(string projectName, string moduleName, string methodName)
        {
            return string.Concat(projectName, ".", moduleName, ".", methodName);
        }
    }
}