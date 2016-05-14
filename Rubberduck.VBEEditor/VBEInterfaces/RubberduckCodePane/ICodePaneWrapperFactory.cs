using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane
{
    public interface ICodePaneWrapperFactory
    {
        ICodePaneWrapper Create(CodePane codePane);
    }
}
