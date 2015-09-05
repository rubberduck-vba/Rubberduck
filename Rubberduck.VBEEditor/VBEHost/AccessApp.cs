using Microsoft.Office.Interop.Access;

namespace Rubberduck.VBEditor.VBEHost
{
    public class AccessApp : HostApplicationBase<Application>
    {
        public AccessApp() : base("Access") { }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            var call = GenerateMethodCall(qualifiedMemberName);
            Application.Run(call);
        }

        private string GenerateMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            //Access only supports Project.Procedure syntax. Error occurs if there are naming conflicts.
            // http://msdn.microsoft.com/en-us/library/office/ff193559(v=office.15).aspx
            // https://github.com/retailcoder/Rubberduck/issues/109

            var projectName = qualifiedMemberName.QualifiedModuleName.Project.Name;
            return string.Concat(projectName, ".", qualifiedMemberName.MemberName);
        }
    }
}