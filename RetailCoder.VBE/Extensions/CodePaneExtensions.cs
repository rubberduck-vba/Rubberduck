using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Extensions
{
    [ComVisible(false)]
    public static class CodePaneExtensions
    {
        public static Selection GetSelection(this CodePane code)
        {
            int startLine;
            int endLine;
            int startColumn;
            int endColumn;

            code.GetSelection(out startLine, out startColumn, out endLine, out endColumn);
            return new Selection(startLine, startColumn, endLine, endColumn);
        }

        public static Selection SelectedProcedure(this CodePane code, Selection selection)
        {
            vbext_ProcKind kind;
            var procedure = code.CodeModule.get_ProcOfLine(selection.StartLine, out kind);
            var startLine = code.CodeModule.get_ProcStartLine(procedure, kind);
            var endLine = startLine + code.CodeModule.get_ProcCountLines(procedure, kind) + 1;

            return new Selection(startLine, 1, endLine, 1);
        }

        /// <summary>
        /// Sets the cursor to the first position of the given line.
        /// </summary>
        /// <param name="codePane"></param>
        /// <param name="lineNumber"></param>
        public static void SetSelection(this CodePane codePane, int lineNumber)
        {
            codePane.SetSelection(lineNumber, 1, lineNumber, 1);
        }
    }
}
