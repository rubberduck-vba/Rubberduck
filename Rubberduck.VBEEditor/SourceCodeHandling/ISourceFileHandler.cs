using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SourceCodeHandling
{
    public interface ISourceFileHandler
    {
        string Export(IVBComponent component);
        void Import(IVBComponent component, string fileName);

        string Read(IVBComponent component);
    }
}
