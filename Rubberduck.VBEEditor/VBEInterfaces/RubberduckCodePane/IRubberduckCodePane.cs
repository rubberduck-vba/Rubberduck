using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane
{
    public interface IRubberduckCodePane
    {
        CodePane CodePane { get; }

        CodePanes Collection { get; }
        VBE VBE { get; }
        Window Window { get; }
        int TopLine { get; set; }
        int CountOfVisibleLines { get; }
        CodeModule CodeModule { get; }
        vbext_CodePaneview CodePaneView { get; }

        void GetSelection(out int startLine, out int startColumn, out int endLine, out int endColumn);
        void SetSelection(int startLine, int startColumn, int endLine, int endColumn);
        void Show();

        /// <summary>   A CodePane extension method that gets the current selection. </summary>
        /// <returns>   The selection. </returns>
        QualifiedSelection GetSelection();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="selection"></param>
        void SetSelection(Selection selection);

        /// <summary>   A CodePane extension method that forces focus onto the CodePane. This patches a bug in VBE.Interop.</summary>
        void ForceFocus();
    }
}