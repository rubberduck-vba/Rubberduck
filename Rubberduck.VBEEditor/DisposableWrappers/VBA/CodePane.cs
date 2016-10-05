using System;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class CodePane : SafeComWrapper<Microsoft.Vbe.Interop.CodePane>, IEquatable<CodePane>
    {
        public CodePane(Microsoft.Vbe.Interop.CodePane codePane)
            : base(codePane)
        {
        }

        public CodePanes Collection
        {
            get { return new CodePanes(InvokeResult(() => IsWrappingNullReference ? null : ComObject.Collection)); }
        }

        public VBE VBE
        {
            get { return new VBE(InvokeResult(() => IsWrappingNullReference ? null : ComObject.VBE)); }
        }

        public Window Window
        {
            get { return new Window(InvokeResult(() => IsWrappingNullReference ? null : ComObject.Window)); }
        }

        public int TopLine
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.TopLine); }
            set { Invoke(() => ComObject.TopLine = value); }
        }

        public int CountOfVisibleLines
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.CountOfVisibleLines); }
        }

        public CodeModule CodeModule
        {
            get { return new CodeModule(InvokeResult(() => IsWrappingNullReference ? null : ComObject.CodeModule)); }
        }

        public CodePaneView CodePaneView
        {
            get { return IsWrappingNullReference ? 0 : (CodePaneView)InvokeResult(() => ComObject.CodePaneView); }
        }

        public Selection GetSelection()
        {
            try
            {
                return InvokeResult(() =>
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
                });
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
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
            Invoke(() => ComObject.SetSelection(startLine, startColumn, endLine, endColumn));
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
            Invoke(() => ComObject.Show());
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