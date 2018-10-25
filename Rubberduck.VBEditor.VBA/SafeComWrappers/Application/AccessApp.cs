using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class AccessApp : HostApplicationBase<Microsoft.Office.Interop.Access.Application>
    {
        public AccessApp() : base("Access") { }

    }
}
