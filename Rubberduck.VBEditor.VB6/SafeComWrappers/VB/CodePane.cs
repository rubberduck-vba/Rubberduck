using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop.VB6;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VB6
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

            var moduleName = QualifiedModuleName;
            return new QualifiedSelection(moduleName, selection);
        }

        public QualifiedModuleName QualifiedModuleName
        {
            get
            {
                using (var codeModule = CodeModule)
                {
                    return codeModule.QualifiedModuleName;
                }
            }
        }

        private void SetSelection(int startLine, int startColumn, int endLine, int endColumn)
        {
            if (IsWrappingNullReference)
            {
                return;
            }
            Target.SetSelection(startLine, startColumn, endLine, endColumn);
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

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}