using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Extensions
{
    [ComVisible(false)]
    public class CodeModuleSelection
    {
        public CodeModuleSelection(CodeModule codeModule, Selection selection)
        {
            _codeModule = codeModule;
            _selection = selection;
        }

        private readonly CodeModule _codeModule;
        public CodeModule CodeModule {get { return _codeModule; } }

        private readonly Selection _selection;
        public Selection Selection { get { return _selection; } }
    }
}