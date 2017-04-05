using System.Globalization;
using System.Windows.Forms;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// A code inspection quickfix that addresses a VBProject bearing the default name.
    /// </summary>
    public class RenameDeclarationQuickFix : IQuickFix
    {
        private readonly Declaration _target;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public RenameDeclarationQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target, RubberduckParserState state, IMessageBox messageBox)
            : base(context, selection, string.Format(RubberduckUI.Rename_DeclarationType, RubberduckUI.ResourceManager.GetString("DeclarationType_" + target.DeclarationType, CultureInfo.CurrentUICulture)))
        {
            _target = target;
            _state = state;
            _messageBox = messageBox;
        }

        public void Fix(IInspectionResult result)
        {
            var vbe = _target.Project.VBE;

            using (var view = new RenameDialog(new RenameViewModel(_state)))
            {
                var factory = new RenamePresenterFactory(vbe, view, _state, _messageBox);
                var refactoring = new RenameRefactoring(vbe, factory, _messageBox, _state);
                refactoring.Refactor(_target);
                IsCancelled = view.DialogResult == DialogResult.Cancel;
            }
        }

        public override bool CanFixInModule { get { return false; } }
        public override bool CanFixInProject { get { return false; } }
    }
}