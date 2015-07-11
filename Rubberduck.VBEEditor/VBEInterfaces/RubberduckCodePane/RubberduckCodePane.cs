using System;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane
{
    public class RubberduckCodePane : IRubberduckCodePane
    {
        private readonly CodePane _codePane;
        public CodePane CodePane { get { return _codePane; } }

        public CodePanes Collection { get { return _codePane.Collection; } }
        public VBE VBE { get { return _codePane.VBE; } }
        public Window Window { get { return _codePane.Window; } }
        public int TopLine { 
            get { return _codePane.TopLine; } 
            set { _codePane.TopLine = value; }
        }
        public int CountOfVisibleLines { get { return _codePane.CountOfVisibleLines; } }
        public CodeModule CodeModule { get { return _codePane.CodeModule; } }
        public vbext_CodePaneview CodePaneView { get { return _codePane.CodePaneView; } }

        public RubberduckCodePane(CodePane codePane)
        {
            _codePane = codePane;
        }

        public void GetSelection(out int startLine, out int startColumn, out int endLine, out int endColumn)
        {
            _codePane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);
        }

        public void SetSelection(int startLine, int startColumn, int endLine, int endColumn)
        {
            _codePane.SetSelection(startLine, startColumn, endLine, endColumn);
        }

        public void Show()
        {
            _codePane.Show();
        }

        /// <summary>   A CodePane extension method that gets the current selection. </summary>
        /// <returns>   The selection. </returns>
        public QualifiedSelection Selection
        {
            get
            {
                int startLine;
                int endLine;
                int startColumn;
                int endColumn;

                if (_codePane == null)
                {
                    return new QualifiedSelection();
                }

                GetSelection(out startLine, out startColumn, out endLine, out endColumn);

                if (endLine > startLine && endColumn == 1)
                {
                    endLine--;
                    endColumn = CodeModule.Lines[endLine, 1].Length;
                }

                var selection = new Selection(startLine, startColumn, endLine, endColumn);
                return new QualifiedSelection(new QualifiedModuleName(CodeModule.Parent), selection);
            }

            set
            {
                var selection = value.Selection;
                SetSelection(selection.StartLine, selection.StartColumn, selection.EndLine, selection.EndColumn);
                ForceFocus();
            }
        }

        /// <summary>   A CodePane extension method that forces focus onto the CodePane. This patches a bug in VBE.Interop.</summary>
        public void ForceFocus()
        {
            Show();

            var mainWindowHandle = VBE.MainWindow.Handle();
            var childWindowFinder = new NativeWindowMethods.ChildWindowFinder(Window.Caption);

            NativeWindowMethods.EnumChildWindows(mainWindowHandle, childWindowFinder.EnumWindowsProcToChildWindowByCaption);
            var handle = childWindowFinder.ResultHandle;

            if (handle != IntPtr.Zero)
            {
                NativeWindowMethods.ActivateWindow(handle, mainWindowHandle);
            }
        }
    }
}