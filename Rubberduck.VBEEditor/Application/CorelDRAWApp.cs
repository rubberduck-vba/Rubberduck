using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Application
{
    public class CorelDRAWApp : HostApplicationBase<Corel.GraphicsSuite.Interop.CorelDRAW.Application>
    {
        public CorelDRAWApp() : base("CorelDRAW") { }
        public CorelDRAWApp(IVBE vbe) : base(vbe, "CorelDRAW") { }

		//Function RunMacro(ModuleName As String, MacroName As String, Parameter() As Variant) As Variant
		//Where ModuleName appears to mean ProjectName, and MacroName has to be provided as "ModuleName.ProcName"

        //TODO:RunMacro can only execute methods in stand-alone projects (not document hosted projects)
        //TODO:Can only get a CorelDraw application if at least one document is open in CorelDraw.

        public override void Run(dynamic declaration)
        {
            var qualifiedMemberName = declaration.QualifiedName;
            var projectName = qualifiedMemberName.QualifiedModuleName.ProjectName;
            var memberName = qualifiedMemberName.QualifiedModuleName.ComponentName + "." + qualifiedMemberName.MemberName;

            if (Application != null)
            {
                Application.GMSManager.RunMacro(projectName, memberName, new object[] {});
            }
        }
    }
}
