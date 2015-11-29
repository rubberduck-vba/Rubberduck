using Microsoft.Vbe.Interop;

namespace Rubberduck.SmartIndenter
{
    public interface IIndenter
    {
        void Indent(VBProject project);
        void Indent(VBComponent module);
        void Indent(VBComponent module, string procedureName);
    }
}
