using System;
using NetOffice.VBIDEApi;

namespace Rubberduck.VBEditor.Extensions
{
    /// <summary>
    /// VBE Code Pane extension methods. 
    /// </summary>
    public static class CodePaneExtensions
    {
        /// <summary>   A CodePane extension method that gets the current selection. </summary>
        /// <returns>   The selection. </returns>
        public static QualifiedSelection GetSelection(this CodePane code)
        {
            int startLine;
            int endLine;
            int startColumn;
            int endColumn;

            if (code == null)
            {
                return new QualifiedSelection();
            }

            code.GetSelection(out startLine, out startColumn, out endLine, out endColumn);

            if (endLine > startLine && endColumn == 1)
            {
                endLine--;
                endColumn = code.CodeModule.get_Lines(endLine, 1).Length;
            }

            var selection = new Selection(startLine, startColumn, endLine, endColumn);
            return new QualifiedSelection(new QualifiedModuleName(code.CodeModule.Parent), selection);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="codePane"></param>
        /// <param name="selection"></param>
        public static void SetSelection(this CodePane codePane, Selection selection)
        {
            codePane.SetSelection(selection.StartLine, selection.StartColumn, selection.EndLine, selection.EndColumn);
            codePane.ForceFocus();
        }

        /// <summary>   A CodePane extension method that forces focus onto the CodePane. This patches a bug in VBE.Interop.</summary>
        public static void ForceFocus(this CodePane codePane)
        {
            codePane.Show();

            var mainWindowHandle =  codePane.VBE.MainWindow.Handle();
            var childWindowFinder = new NativeWindowMethods.ChildWindowFinder(codePane.Window.Caption);

            NativeWindowMethods.EnumChildWindows(mainWindowHandle, childWindowFinder.EnumWindowsProcToChildWindowByCaption);
            var handle = childWindowFinder.ResultHandle;

            if (handle != IntPtr.Zero)
            {
                NativeWindowMethods.ActivateWindow(handle, mainWindowHandle);
            }
        }
    }
}
