using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane
{
    public class RubberduckCodePaneFactory : IRubberduckCodePaneFactory
    {
        public IRubberduckCodePane Create(CodePane codePane)
        {
            return new RubberduckCodePane(codePane);
        }
    }
}
