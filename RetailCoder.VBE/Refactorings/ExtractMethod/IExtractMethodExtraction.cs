using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodeModule;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractMethodExtraction
    {
        void apply(ICodeModuleWrapper codeModule, IExtractMethodModel model, Selection selection);
    }
}
