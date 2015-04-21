using Microsoft.Office.Interop.Excel;
using Rubberduck.Parsing;

namespace Rubberduck.VBEHost
{
    public class PowerPointApp : HostApplicationBase<Application>
    {
        public PowerPointApp() : base("PowerPoint") { }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            object[] paramArray = { }; //PowerPoint requires a paramarray, so we pass it an empty array.

            var call = GenerateMethodCall(qualifiedMemberName);
            Application.Run(call, paramArray);
        }

        protected virtual string GenerateMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            /* Note: Powerpoint supports a `FileName.ppt!Module.method` syntax
             * http://msdn.microsoft.com/en-us/library/office/ff744221(v=office.15).aspx
             */

            var projectFile = qualifiedMemberName.QualifiedModuleName.Project.FileName;
            var moduleName = qualifiedMemberName.QualifiedModuleName.Component.Name;

            return string.Concat(projectFile, "!", moduleName, ".", qualifiedMemberName.MemberName);
        }
    }
}