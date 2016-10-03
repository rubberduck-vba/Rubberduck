using System;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.DisposableWrappers;
using Rubberduck.VBEditor.Extensions;

public static class CodePaneExtensions
{
    public static QualifiedSelection? GetQualifiedSelection(this CodePane pane)
    {
        if (pane.IsWrappingNullReference)
        {
            return null;
        }

        var selection = pane.GetSelection();
        VBComponent component = null;
        if (selection.EndLine > selection.StartLine && selection.EndColumn == 1)
        {
            var endLine = selection.EndLine - 1;
            using (var module = pane.CodeModule)
            {
                var endColumn = module.GetLines(endLine, 1).Length;
                selection = new Selection(selection.StartLine, selection.StartColumn, selection.EndLine, endColumn);
                component = module.Parent;
            }
        }

        var moduleName = new QualifiedModuleName(component);
        return new QualifiedSelection(moduleName, selection);
    }

    public static Selection? GetSelection(this CodePane pane)
    {
        if (pane == null)
        {
            return null;
        }

        var selection = pane.GetSelection();
        if (selection.EndLine > selection.StartLine && selection.EndColumn == 1)
        {
            var endLine = selection.EndLine - 1;
            using (var module = pane.CodeModule)
            {
                var endColumn = module.GetLines(endLine, 1).Length;
                selection = new Selection(selection.StartLine, selection.StartColumn, selection.EndLine, endColumn);
            }
        }

        return selection;
    }

    public static void ForceFocus(this CodePane pane)
    {
        pane.Show();
        IntPtr mainWindowHandle;
        string caption;

        using (var vbe = pane.VBE)
        using (var window = vbe.MainWindow)
        {
            mainWindowHandle = window.Handle();
            caption = window.Caption;
        }

        var childWindowFinder = new NativeMethods.ChildWindowFinder(caption);

        NativeMethods.EnumChildWindows(mainWindowHandle, childWindowFinder.EnumWindowsProcToChildWindowByCaption);
        var handle = childWindowFinder.ResultHandle;

        if (handle != IntPtr.Zero)
        {
            NativeMethods.ActivateWindow(handle, mainWindowHandle);
        }
    }

}