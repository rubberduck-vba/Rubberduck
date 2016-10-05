using Interop.SldWorks.Types;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

namespace Rubberduck.VBEditor.VBEHost
{
    public class SolidWorksApp : HostApplicationBase<Interop.SldWorks.Extensibility.Application>
    {
        public SolidWorksApp() : base("SolidWorks") { }
        public SolidWorksApp(VBE vbe) : base(vbe, "SolidWorks") { }
		
        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            var projectFileName = qualifiedMemberName.QualifiedModuleName.Project.FileName;
            var moduleName = qualifiedMemberName.QualifiedModuleName.ComponentName;
            var memberName = qualifiedMemberName.MemberName;

            if (Application != null)
            {
                SldWorks runner = (SldWorks)Application.SldWorks;
                runner.RunMacro(projectFileName, moduleName, memberName);
            }
        }
    }
}
