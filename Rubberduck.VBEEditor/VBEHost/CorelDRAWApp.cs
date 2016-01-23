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
            //object oMissing = System.Reflection.Missing.Value;
            Application.GMSManager.RunMacro(qualifiedMemberName.QualifiedModuleName.ProjectName.ToString(), qualifiedMemberName.ToString(), null);
        }
    }
}