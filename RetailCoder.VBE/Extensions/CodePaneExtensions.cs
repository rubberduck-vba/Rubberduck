using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using System;
using Rubberduck.UI;

namespace Rubberduck.Extensions
{
    /// <summary>   VBE Code Pane extensions. </summary>
    [ComVisible(false)]
    public static class CodePaneExtensions
    {
        /// <summary>   A CodePane extension method that gets the current selection. </summary>
        /// <returns>   The selection. </returns>
        public static Selection GetSelection(this CodePane code)
        {
            int startLine;
            int endLine;
            int startColumn;
            int endColumn;

            code.GetSelection(out startLine, out startColumn, out endLine, out endColumn);
            return new Selection(startLine, startColumn, endLine, endColumn);
        }

        /// <summary>   A CodePane extension method that selected procedure. </summary>
        ///
        /// <param name="selection">    The selection. </param>
        /// <returns>   A Selection object representing the procedure the cursor is currently in. </returns>
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

        /// <summary>   A CodePane extension method that forces focus onto the CodePane. This patches a bug in VBE.Interop.</summary>
        public static void ForceFocus(this CodePane codePane)
        {
            codePane.Show();

            IntPtr mainWindowHandle =  codePane.VBE.MainWindow.Handle();
            var childWindowFinder = new NativeWindowMethods.ChildWindowFinder(codePane.Window.Caption);

            NativeWindowMethods.EnumChildWindows(mainWindowHandle, childWindowFinder.EnumWindowsProcToChildWindowByCaption);
            IntPtr handle = childWindowFinder.ResultHandle;

            if (handle != IntPtr.Zero)
            {
                NativeWindowMethods.ActivateWindow(handle, mainWindowHandle);
            }

        }
    }
}
