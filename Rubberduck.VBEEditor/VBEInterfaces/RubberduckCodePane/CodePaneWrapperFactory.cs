using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane
{
    public class CodePaneWrapperFactory : ICodePaneWrapperFactory
    {
        public ICodePaneWrapper Create(CodePane codePane)
        {
            return new CodePaneWrapper(codePane);
        }
    }
}
