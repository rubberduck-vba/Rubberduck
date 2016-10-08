using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Application
{
    public class ProjectApp : HostApplicationBase<Microsoft.Office.Interop.MSProject.Application>
    {
        public ProjectApp() : base("MSProject") { }
        public ProjectApp(IVBE vbe) : base(vbe, "MSProject") { }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            var call = GenerateMethodCall(qualifiedMemberName);
            Application.Macro(call);
        }

        protected virtual string GenerateMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            var moduleName = qualifiedMemberName.QualifiedModuleName.Component.Name;
            return string.Concat(moduleName, ".", qualifiedMemberName.MemberName);
        }
    }
}
