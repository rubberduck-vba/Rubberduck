using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SourceCodeHandling
{
    public interface IComponentSourceCodeHandler
    {
        string SourceCode(IVBComponent module);
        IVBComponent SubstituteCode(IVBComponent module, string newCode);
    }
}