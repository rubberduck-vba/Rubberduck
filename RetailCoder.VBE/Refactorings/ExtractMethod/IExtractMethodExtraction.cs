using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractMethodExtraction
    {
        void Apply(ICodeModule codeModule, IExtractMethodModel model, Selection selection);
    }
}
