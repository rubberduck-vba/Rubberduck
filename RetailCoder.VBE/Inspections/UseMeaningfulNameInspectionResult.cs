using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public class UseMeaningfulNameInspectionResult : CodeInspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public UseMeaningfulNameInspectionResult(IInspection inspection, Declaration target, RubberduckParserState parserState, ICodePaneWrapperFactory wrapperFactory, IMessageBox messageBox)
            : base(inspection, string.Format(inspection.Description, RubberduckUI.ResourceManager.GetString("DeclarationType_" + target.DeclarationType, RubberduckUI.Culture), target.IdentifierName), target)
        {
            _quickFixes = new[]
            {
                new RenameDeclarationQuickFix(target.Context, target.QualifiedSelection, target, parserState, wrapperFactory, messageBox),
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    /// <summary>
    /// A code inspection quickfix that addresses a VBProject bearing the default name.
    /// </summary>
    public class RenameDeclarationQuickFix : CodeInspectionQuickFix
    {
        private readonly Declaration _target;
        private readonly RubberduckParserState _state;
        private readonly ICodePaneWrapperFactory _wrapperFactory;
        private readonly IMessageBox _messageBox;

        public RenameDeclarationQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target, RubberduckParserState state, ICodePaneWrapperFactory wrapperFactory, IMessageBox messageBox)
            : base(context, selection, string.Format(RubberduckUI.Rename_DeclarationType, RubberduckUI.ResourceManager.GetString("DeclarationType_" + target.DeclarationType, RubberduckUI.Culture)))
        {
            _target = target;
            _state = state;
            _wrapperFactory = wrapperFactory;
            _messageBox = messageBox;
        }

        public override void Fix()
        {
            var vbe = Selection.QualifiedName.Project.VBE;

            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(vbe, view, _state, _messageBox, _wrapperFactory);
                var refactoring = new RenameRefactoring(factory, new ActiveCodePaneEditor(vbe, _wrapperFactory), _messageBox, _state);
                refactoring.Refactor(_target);
            }
        }

        public override bool CanFixInModule { get { return false; } }
        public override bool CanFixInProject { get { return false; } }
    }
}