using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.WindowsApi;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class CodePane : SafeComWrapper<VB.CodePane>, ICodePane
    {
        public CodePane(VB.CodePane target, bool rewrapping = false)
            : base(target, rewrapping)
        {
        }

        public ICodePanes Collection => new CodePanes(IsWrappingNullReference ? null : Target.Collection);

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        public IWindow Window => new Window(IsWrappingNullReference ? null : Target.Window);

        public int TopLine
        {
            get => IsWrappingNullReference ? 0 : Target.TopLine;
            set { if (!IsWrappingNullReference) Target.TopLine = value; }
        }

        public int CountOfVisibleLines => IsWrappingNullReference ? 0 : Target.CountOfVisibleLines;

        public ICodeModule CodeModule => new CodeModule(IsWrappingNullReference ? null : Target.CodeModule);

        public CodePaneView CodePaneView => IsWrappingNullReference ? 0 : (CodePaneView)Target.CodePaneView;

        public Selection Selection
        {
            get => GetSelection();
            set { if (!IsWrappingNullReference) SetSelection(value.StartLine, value.StartColumn, value.EndLine, value.EndColumn); }
        }

        private Selection GetSelection()
        {
            if (IsWrappingNullReference) return new Selection(0, 0, 0, 0);

            Target.GetSelection(out var startLine, out var startColumn, out var endLine, out var endColumn);

            if (endLine > startLine && endColumn == 1)
            {
                endLine -= 1;
                using (var codeModule = CodeModule)
                {
                    endColumn = codeModule.GetLines(endLine, 1).Length;
                }
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

            IVBComponent component;
            using (var codeModule = CodeModule)
            {
                component = new VBComponent((VB.VBComponent)codeModule.Parent.Target, rewrapping: true);
            }
            var moduleName = new QualifiedModuleName(component);
            return new QualifiedSelection(moduleName, selection);
        }

        private void SetSelection(int startLine, int startColumn, int endLine, int endColumn)
        {
            if (IsWrappingNullReference)
            {
                return;
            }
            Target.SetSelection(startLine, startColumn, endLine, endColumn);
            ForceFocus();
        }

        private void ForceFocus()
        {
            if (IsWrappingNullReference) return;
            Show();

            var window = VBE.MainWindow;
            var mainWindowHandle = window.Handle();
            var caption = Window.Caption;
            var childWindowFinder = new ChildWindowFinder(caption);

            NativeMethods.EnumChildWindows(mainWindowHandle, childWindowFinder.EnumWindowsProcToChildWindowByCaption);
            var handle = childWindowFinder.ResultHandle;

            if (handle != IntPtr.Zero)
            {
                NativeMethods.ActivateWindow(handle, mainWindowHandle);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("CodePane.ForceFocus() failed to get a handle on the MainWindow.");
            }
        }

        public void Show()
        {
            if (IsWrappingNullReference) return;
            Target.Show();
        }

        public override bool Equals(ISafeComWrapper<VB.CodePane> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(ICodePane other)
        {
            return Equals(other as SafeComWrapper<VB.CodePane>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}