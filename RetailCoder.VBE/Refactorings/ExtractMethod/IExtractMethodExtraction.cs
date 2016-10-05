using Rubberduck.VBEditor;
using Rubberduck.VBEditor.DisposableWrappers;
using Rubberduck.VBEditor.DisposableWrappers.VBA;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractMethodExtraction
    {
        void Apply(CodeModule codeModule, IExtractMethodModel model, Selection selection);
    }
}
