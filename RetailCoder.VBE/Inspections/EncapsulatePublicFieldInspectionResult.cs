using System.Collections.Generic;
using System.Windows.Forms;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public class EncapsulatePublicFieldInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public EncapsulatePublicFieldInspectionResult(IInspection inspection, Declaration target, RubberduckParserState state, ICodePaneWrapperFactory wrapperFactory)
            : base(inspection, target)
        {
            _quickFixes = new[]
            {
                new EncapsulateFieldQuickFix(target.Context, target.QualifiedSelection, target, state, wrapperFactory),
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.EncapsulatePublicFieldInspectionResultFormat, Target.IdentifierName); }
        }
    }

    /// <summary>
    /// A code inspection quickfix that encapsulates a public field with a property
    /// </summary>
    public class EncapsulateFieldQuickFix : CodeInspectionQuickFix
    {
        private readonly Declaration _target;
        private readonly RubberduckParserState _state;
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public EncapsulateFieldQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target, RubberduckParserState state, ICodePaneWrapperFactory wrapperFactory)
            : base(context, selection, string.Format(InspectionsUI.EncapsulatePublicFieldInspectionQuickFix, target.IdentifierName))
        {
            _target = target;
            _state = state;
            _wrapperFactory = wrapperFactory;
        }

        public override void Fix()
        {
            var vbe = Selection.QualifiedName.Project.VBE;

            using (var view = new EncapsulateFieldDialog())
            {
                var factory = new EncapsulateFieldPresenterFactory(vbe, _state, view);
                var refactoring = new EncapsulateFieldRefactoring(vbe, factory);
                refactoring.Refactor(_target);
                IsCancelled = view.DialogResult != DialogResult.OK;
            }
        }

        public override bool CanFixInModule { get { return false; } }
        public override bool CanFixInProject { get { return false; } }
    }
}
