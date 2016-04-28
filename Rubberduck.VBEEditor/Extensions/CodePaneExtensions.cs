using System;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

public static class CodePaneExtensions
{
    public static QualifiedSelection? GetQualifiedSelection(this CodePane pane)
    {
        int startLine;
        int endLine;
        int startColumn;
        int endColumn;

        if (pane == null)
        {
            return null;
        }

        pane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);

        if (endLine > startLine && endColumn == 1)
        {
            endLine--;
            endColumn = pane.CodeModule.get_Lines(endLine, 1).Length;
        }

        var selection = new Selection(startLine, startColumn, endLine, endColumn);
        var moduleName = new QualifiedModuleName(pane.CodeModule.Parent);
        return new QualifiedSelection(moduleName, selection);
    }

    public static Selection GetSelection(this CodePane pane)
    {
        int startLine;
        int endLine;
        int startColumn;
        int endColumn;

        pane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);

        if (endLine > startLine && endColumn == 1)
        {
            endLine--;
            endColumn = pane.CodeModule.Lines[endLine, 1].Length;
        }

        return new Selection(startLine, startColumn, endLine, endColumn);
    }

    public static void ForceFocus(this CodePane pane)
    {
        pane.Show();

        var mainWindowHandle = pane.VBE.MainWindow.Handle();
        var childWindowFinder = new NativeMethods.ChildWindowFinder(pane.Window.Caption);

        NativeMethods.EnumChildWindows(mainWindowHandle, childWindowFinder.EnumWindowsProcToChildWindowByCaption);
        var handle = childWindowFinder.ResultHandle;

        if (handle != IntPtr.Zero)
        {
            NativeMethods.ActivateWindow(handle, mainWindowHandle);
        }
    }

}