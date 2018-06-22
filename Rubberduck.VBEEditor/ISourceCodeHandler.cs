using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor
{
    public interface ISourceCodeHandler
    {
        string Export(IVBComponent component);
        void Import(IVBComponent component, string fileName);

        string Read(IVBComponent component);
    }
}
