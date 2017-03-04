using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Application
{
    public class ProjectApp : HostApplicationBase<Microsoft.Office.Interop.MSProject.Application>
    {
        public ProjectApp() : base("MSProject") { }
        public ProjectApp(IVBE vbe) : base(vbe, "MSProject") { }

        public override void Run(dynamic declaration)
        {
            var call = GenerateMethodCall(declaration.QualifiedName);
            Application.Macro(call);
        }

        protected virtual string GenerateMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            var moduleName = qualifiedMemberName.QualifiedModuleName.ComponentName;
            return string.Concat(moduleName, ".", qualifiedMemberName.MemberName);
        }
    }
}
