using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.VBA;
using Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractMethodExtraction
    {
        void Apply(ICodeModule codeModule, IExtractMethodModel model, Selection selection);
    }
}
