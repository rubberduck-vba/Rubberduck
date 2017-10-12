using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractMethodExtraction
    {
        void Apply(ICodeModule codeModule, IExtractMethodModel model, Selection selection);
    }
}
