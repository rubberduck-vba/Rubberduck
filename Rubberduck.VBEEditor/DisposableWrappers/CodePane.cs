using System;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class CodePane : WrapperBase<Microsoft.Vbe.Interop.CodePane>
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
                int startLine;
                int startColumn;
                int endLine;
                int endColumn;
                Item.GetSelection(out startLine, out startColumn, out endLine, out endColumn);
                return new Selection(startLine, startColumn, endLine, endColumn);
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        public void SetSelection(int startLine, int startColumn, int endLine, int endColumn)
        {
            ThrowIfDisposed();
            InvokeMember((startL, startC, endL, endC) => Item.SetSelection(startL, startC, endL, endC), startLine, startColumn, endLine, endColumn);
        }

        public void SetSelection(Selection selection)
        {
            SetSelection(selection.StartLine, selection.StartColumn, selection.EndLine, selection.EndColumn);
        }

        public void Show()
        {
            ThrowIfDisposed();
            InvokeMember(() => Item.Show());
        }

        public CodePanes Collection
        {
            get
            {
                ThrowIfDisposed();
                throw new NotImplementedException();
            }
        }

        public VBE VBE
        {
            get
            {
                ThrowIfDisposed();
                return new VBE(InvokeMemberValue(() => Item.VBE));
            }
        }

        public Window Window
        {
            get
            {
                ThrowIfDisposed();
                return new Window(InvokeMemberValue(() => Item.Window));
            }
        }

        public int TopLine
        {
            get
            {
                ThrowIfDisposed();
                return InvokeMemberValue(() => Item.TopLine);
            }
            set
            {
                ThrowIfDisposed();
                Item.TopLine = value;
            }
        }

        public int CountOfVisibleLines
        {
            get
            {
                ThrowIfDisposed();
                return InvokeMemberValue(() => Item.CountOfVisibleLines);
            }
        }

        public CodeModule CodeModule
        {
            get
            {
                ThrowIfDisposed();
                throw new NotImplementedException();
            }
        }

        public vbext_CodePaneview CodePaneView
        {
            get
            {
                ThrowIfDisposed();
                throw new NotImplementedException();
            }
        }
    }
}