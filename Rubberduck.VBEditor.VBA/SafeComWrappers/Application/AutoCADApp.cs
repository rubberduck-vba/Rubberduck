using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class AutoCADApp : HostApplicationBase<Autodesk.AutoCAD.Interop.AcadApplication>
    {
        public AutoCADApp() : base("AutoCAD") { }

    }
}
