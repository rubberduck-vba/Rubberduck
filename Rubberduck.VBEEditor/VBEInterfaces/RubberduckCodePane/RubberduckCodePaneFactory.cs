using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane
{
    public class RubberduckCodePaneFactory : IRubberduckFactory<IRubberduckCodePane>
    {
        public IRubberduckCodePane Create(object codePane)
        {
            return new RubberduckCodePane(codePane as CodePane);
        }
    }
}
