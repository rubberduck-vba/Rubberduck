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
    public class DefaultProjectNameInspectionResult : CodeInspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes; 

        public DefaultProjectNameInspectionResult(IInspection inspection, Declaration target, IRubberduckParserState parseResult, ICodePaneWrapperFactory wrapperFactory)
            : base(inspection, string.Format(inspection.Description, target.IdentifierName), target)
        {
            _quickFixes = new[]
            {
                new RenameProjectQuickFix(target.Context, target.QualifiedSelection, target, parseResult, wrapperFactory),
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    /// <summary>
    /// A code inspection quickfix that addresses a VBProject bearing the default name.
    /// </summary>
    public class RenameProjectQuickFix : CodeInspectionQuickFix
    {
        private readonly Declaration _target;
        private readonly IRubberduckParserState _parseResult;
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public RenameProjectQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target, IRubberduckParserState parseResult, ICodePaneWrapperFactory wrapperFactory)
            : base(context, selection, string.Format(RubberduckUI.Rename_DeclarationType, RubberduckUI.ResourceManager.GetString("DeclarationType_" + DeclarationType.Project, RubberduckUI.Culture)))
        {
            _target = target;
            _parseResult = parseResult;
            _wrapperFactory = wrapperFactory;
        }

        public override void Fix()
        {
            var vbe = Selection.QualifiedName.Project.VBE;

            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(vbe, view, _parseResult, new MessageBox(), _wrapperFactory);
                var refactoring = new RenameRefactoring(factory, new ActiveCodePaneEditor(vbe, _wrapperFactory), new MessageBox());
                refactoring.Refactor(_target);
            }
        }

        public override bool CanFixInModule { get { return false; } }
        public override bool CanFixInProject { get { return false; } }
    }
}
