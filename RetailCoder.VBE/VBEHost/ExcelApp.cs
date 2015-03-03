using Microsoft.Office.Interop.Excel;

namespace Rubberduck.VBEHost
{
    public class ExcelApp : HostApplicationBase<Application>
    {
        public ExcelApp() : base("Excel") { }

        public override void Run(string target)
        {
            base.Application.Run(target);
        }

        protected override string GenerateFullyQualifiedName(string projectName, string moduleName, string methodName)
        {
            return string.Concat(projectName, ".", moduleName, ".", methodName);
        }
    }
}