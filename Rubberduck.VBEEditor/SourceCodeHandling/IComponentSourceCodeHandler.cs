using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SourceCodeHandling
{
    public interface IComponentSourceCodeHandler
    {
        string SourceCode(IVBComponent module);
        void SubstituteCode(IVBComponent module, string newCode);
    }
}