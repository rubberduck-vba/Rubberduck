using Corel.GraphicsSuite.Interop.CorelDRAW;

namespace Rubberduck.VBEditor.VBEHost
{
    public class CorelDRAWApp : HostApplicationBase<Application>
    {
        public CorelDRAWApp() : base("CorelDRAW") { }

		//Function RunMacro(ModuleName As String, MacroName As String, Parameter() As Variant) As Variant
		//Where ModuleName appears to mean ProjectName, and MacroName has to be provided as "ModuleName.ProcName"
		
        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            var projectName = qualifiedMemberName.QualifiedModuleName.ProjectName;
            var memberName = qualifiedMemberName.QualifiedModuleName.ComponentName + "." + qualifiedMemberName.MemberName;

            RunHelper(projectName, memberName);
        }

        private void RunHelper(string ProjectName, string MemberName, params object[] p)
        {
            //Application.GMSManager.RunMacro(ProjectName, MemberName, p);
            //object GMS = Application.GMSManager;
        }
    }
}