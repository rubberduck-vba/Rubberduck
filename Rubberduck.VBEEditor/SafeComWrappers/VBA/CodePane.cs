using System;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class CodePane : SafeComWrapper<Microsoft.Vbe.Interop.CodePane>, IEquatable<CodePane>
    {
        public CodePane(Microsoft.Vbe.Interop.CodePane codePane)
            : base(codePane)
        {
        }

        public CodePanes Collection
        {
            get { return new CodePanes(IsWrappingNullReference ? null : ComObject.Collection); }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : ComObject.VBE); }
        }

        public IWindow Window
        {
            get { return new Window(IsWrappingNullReference ? null : ComObject.Window); }
        }

        public int TopLine
        {
            get { return IsWrappingNullReference ? 0 : ComObject.TopLine; }
            set { ComObject.TopLine = value; }
        }

        public int CountOfVisibleLines
        {
            get { return IsWrappingNullReference ? 0 : ComObject.CountOfVisibleLines; }
        }
        
        public ICodeModule CodeModule
        {
            get { return new CodeModule(IsWrappingNullReference ? null : ComObject.CodeModule); }
        }

        public CodePaneView CodePaneView
        {
            get { return IsWrappingNullReference ? 0 : (CodePaneView)ComObject.CodePaneView; }
        }

        public Selection GetSelection()
        {
            int startLine;
            int startColumn;
            int endLine;
            int endColumn;
            ComObject.GetSelection(out startLine, out startColumn, out endLine, out endColumn);

            if (endLine > startLine && endColumn == 1)
            {
                endLine -= 1;
                endColumn = CodeModule.GetLines(endLine, 1).Length;
            }

            return new Selection(startLine, startColumn, endLine, endColumn);
        }

        public QualifiedSelection? GetQualifiedSelection()
        {
            if (IsWrappingNullReference)
            {
                return null;
            }

            var selection = GetSelection();
            if (selection.IsEmpty())
            {
                return null;
            }

            var component = new VBComponent(CodeModule.Parent.ComObject);
            var moduleName = new QualifiedModuleName(component);
            return new QualifiedSelection(moduleName, selection);
        }

        public void SetSelection(int startLine, int startColumn, int endLine, int endColumn)
        {
            ComObject.SetSelection(startLine, startColumn, endLine, endColumn);
            ForceFocus();
        }

        public void SetSelection(Selection selection)
        {
            SetSelection(selection.StartLine, selection.StartColumn, selection.EndLine, selection.EndColumn);
        }

        private void ForceFocus()
        {
            Show();

            var window = VBE.MainWindow;
            var mainWindowHandle = window.Handle();
            var caption = window.Caption;
            var childWindowFinder = new NativeMethods.ChildWindowFinder(caption);

            NativeMethods.EnumChildWindows(mainWindowHandle, childWindowFinder.EnumWindowsProcToChildWindowByCaption);
            var handle = childWindowFinder.ResultHandle;

            if (handle != IntPtr.Zero)
            {
                NativeMethods.ActivateWindow(handle, mainWindowHandle);
            }
        }

        public void Show()
        {
            ComObject.Show();
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                //Window.Release(); window is released by VBE.Windows
                Marshal.ReleaseComObject(ComObject);
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.CodePane> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.ComObject, ComObject));
        }

        public bool Equals(CodePane other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.CodePane>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}