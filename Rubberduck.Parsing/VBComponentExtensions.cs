using Microsoft.Vbe.Interop;

namespace Rubberduck.Parsing
{
    public static class VBComponentExtensions
    {
        public static QualifiedModuleName QualifiedName(this VBComponent component)
        {
            return new QualifiedModuleName(component);
        }
    }
}
