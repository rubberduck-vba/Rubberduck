using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane
{
    public interface IRubberduckCodePaneFactory
    {
        IRubberduckCodePane Create(CodePane codePane);
    }
}