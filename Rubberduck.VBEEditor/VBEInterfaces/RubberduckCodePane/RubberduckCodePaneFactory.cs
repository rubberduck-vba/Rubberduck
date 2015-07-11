using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane
{
    public class RubberduckCodePaneFactory : IRubberduckFactory<IRubberduckCodePane>
    {
        private readonly CodePane _codePane;

        public RubberduckCodePaneFactory(CodePane codePane)
        {
            _codePane = codePane;
        }

        public IRubberduckCodePane Create()
        {
            return new RubberduckCodePane(_codePane);
        }
    }
}
