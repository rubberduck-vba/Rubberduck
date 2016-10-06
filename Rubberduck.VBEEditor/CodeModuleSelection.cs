using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor
{
    public class CodeModuleSelection
    {
        public CodeModuleSelection(ICodeModule codeModule, Selection selection)
        {
            _codeModule = codeModule;
            _selection = selection;
        }

        private readonly ICodeModule _codeModule;
        public ICodeModule CodeModule {get { return _codeModule; } }

        private readonly Selection _selection;
        public Selection Selection { get { return _selection; } }
    }
}
