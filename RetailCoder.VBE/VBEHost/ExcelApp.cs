using Microsoft.Office.Interop.Excel;

namespace Rubberduck.VBEHost
{
    public class ExcelApp : HostApplicationBase<Application>
    {
        public ExcelApp() : base("Excel") { }

        public override void Run(string projectName, string moduleName, string methodName)
        {
            Application.Run(GenerateMethodCall(projectName, moduleName, methodName));
        }

        protected override string GenerateMethodCall(string projectName, string moduleName, string methodName)
        {
            return string.Concat(projectName, ".", moduleName, ".", methodName);
        }
    }
}