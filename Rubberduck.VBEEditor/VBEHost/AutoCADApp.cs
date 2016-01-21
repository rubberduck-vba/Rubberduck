using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;

namespace Rubberduck.VBEditor.VBEHost
{
    public class AutoCADApp : HostApplicationBase<Autodesk.AutoCAD.Interop.AcadApplication>
    {
        public AutoCADApp() : base("AutoCAD") { }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            Application.RunMacro(qualifiedMemberName.ToString());
        }
    }
}