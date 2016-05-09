using System.Collections.Generic;
using System.Windows.Forms;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using MessageBox = Rubberduck.UI.MessageBox;

namespace Rubberduck.Inspections
{
    public class DefaultProjectNameInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes; 

        public DefaultProjectNameInspectionResult(IInspection inspection, Declaration target, RubberduckParserState parseResult, ICodePaneWrapperFactory wrapperFactory)
            : base(inspection, target)
        {
            _quickFixes = new[]
            {
                new RenameProjectQuickFix(target.Context, target.QualifiedSelection, target, parseResult, wrapperFactory),
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return Inspection.Description; }
        }
    }

    /// <summary>
    /// A code inspection quickfix that addresses a VBProject bearing the default name.
    /// </summary>
    public class RenameProjectQuickFix : CodeInspectionQuickFix
    {
        private readonly Declaration _target;
        private readonly RubberduckParserState _state;
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public RenameProjectQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target, RubberduckParserState state, ICodePaneWrapperFactory wrapperFactory)
            : base(context, selection, string.Format(RubberduckUI.Rename_DeclarationType, RubberduckUI.ResourceManager.GetString("DeclarationType_" + DeclarationType.Project, RubberduckUI.Culture)))
        {
            _target = target;
            _state = state;
            _wrapperFactory = wrapperFactory;
        }

        public override void Fix()
        {
            var vbe = Selection.QualifiedName.Project.VBE;

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
