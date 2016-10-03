using Rubberduck.VBEditor;
using Rubberduck.VBEditor.DisposableWrappers;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractMethodExtraction
    {
        void Apply(CodeModule codeModule, IExtractMethodModel model, Selection selection);
    }
}
