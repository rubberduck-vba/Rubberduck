using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane
{
    public interface ICodePaneWrapper
    {
        CodePane CodePane { get; }

        CodePanes Collection { get; }
        VBE VBE { get; }
        Window Window { get; }
        int TopLine { get; set; }
        int CountOfVisibleLines { get; }
        CodeModule CodeModule { get; }
        vbext_CodePaneview CodePaneView { get; }
        Selection Selection { get; set; }

        void GetSelection(out int startLine, out int startColumn, out int endLine, out int endColumn);
        void SetSelection(int startLine, int startColumn, int endLine, int endColumn);
        void Show();

        /// <summary>   A CodePane extension method that forces focus onto the CodePane. This patches a bug in VBE.Interop.</summary>
        void ForceFocus();
    }
}