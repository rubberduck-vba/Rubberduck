using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class CodePane : SafeComWrapper<Microsoft.Vbe.Interop.CodePane>
    {
        public CodePane(Microsoft.Vbe.Interop.CodePane codePane)
            : base(codePane)
        {
        }

        public Selection GetSelection()
        {
            ThrowIfDisposed();
            try
            {
                return InvokeResult(() =>
                {
                    int startLine;
                    int startColumn;
                    int endLine;
                    int endColumn;
                    ComObject.GetSelection(out startLine, out startColumn, out endLine, out endColumn);
                    return new Selection(startLine, startColumn, endLine, endColumn);
                });
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        public void SetSelection(int startLine, int startColumn, int endLine, int endColumn)
        {
            Invoke(() => ComObject.SetSelection(startLine, startColumn, endLine, endColumn));
        }

        public void SetSelection(Selection selection)
        {
            Invoke(() => ComObject.SetSelection(selection.StartLine, selection.StartColumn, selection.EndLine, selection.EndColumn));
        }

        public void Show()
        {
            Invoke(() => ComObject.Show());
        }

        public CodePanes Collection { get { return new CodePanes(InvokeResult(() => ComObject.Collection)); } }

        public VBE VBE { get { return new VBE(InvokeResult(() => ComObject.VBE)); } }

        public Window Window { get { return new Window(InvokeResult(() => ComObject.Window)); } }

        public int TopLine 
        { 
            get { return InvokeResult(() => ComObject.TopLine); }
            set { Invoke(() => ComObject.TopLine = value); }
        }

        public int CountOfVisibleLines { get { return InvokeResult(() => ComObject.CountOfVisibleLines); } }

        public CodeModule CodeModule { get { return new CodeModule(InvokeResult(() => ComObject.CodeModule)); } }

        public CodePaneView CodePaneView { get { return (CodePaneView)InvokeResult(() => ComObject.CodePaneView); } }
    }
}