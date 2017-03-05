using System.Globalization;
using System.Windows.Forms;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;
using MessageBox = Rubberduck.UI.MessageBox;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// A code inspection quickfix that addresses a VBProject bearing the default name.
    /// </summary>
    public class RenameProjectQuickFix : QuickFixBase
    {
        private readonly Declaration _target;
        private readonly RubberduckParserState _state;

        public RenameProjectQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target, RubberduckParserState state)
            : base(context, selection, string.Format(RubberduckUI.Rename_DeclarationType, RubberduckUI.ResourceManager.GetString("DeclarationType_" + DeclarationType.Project, CultureInfo.CurrentUICulture)))
        {
            _target = target;
            _state = state;
        }

        public override void Fix()
        {
            var vbe = _target.Project.VBE;

            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(vbe, view, _state, new MessageBox());
                var refactoring = new RenameRefactoring(vbe, factory, new MessageBox(), _state);
                refactoring.Refactor(_target);
                IsCancelled = view.DialogResult == DialogResult.Cancel;
            }
        }

        public override bool CanFixInModule { get { return false; } }
        public override bool CanFixInProject { get { return false; } }
    }
}