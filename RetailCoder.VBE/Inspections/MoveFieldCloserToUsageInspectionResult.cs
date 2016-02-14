using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public class MoveFieldCloserToUsageInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public MoveFieldCloserToUsageInspectionResult(IInspection inspection, Declaration target, RubberduckParserState parseResult, ICodePaneWrapperFactory wrapperFactory, IMessageBox messageBox)
            : base(inspection, string.Format(inspection.Description, target.IdentifierName), target)
        {
            _quickFixes = new[]
            {
                new MoveFieldCloserToUsageQuickFix(target.Context, target.QualifiedSelection, target, parseResult, wrapperFactory, messageBox),
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    /// <summary>
    /// A code inspection quickfix that encapsulates a public field with a property
    /// </summary>
    public class MoveFieldCloserToUsageQuickFix : CodeInspectionQuickFix
    {
        private readonly Declaration _target;
        private readonly RubberduckParserState _parseResult;
        private readonly ICodePaneWrapperFactory _wrapperFactory;
        private readonly IMessageBox _messageBox;

        public MoveFieldCloserToUsageQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target, RubberduckParserState parseResult, ICodePaneWrapperFactory wrapperFactory, IMessageBox messageBox)
            : base(context, selection, string.Format(InspectionsUI.MoveFieldCloserToUsageInspectionResultFormat, target.IdentifierName))
        {
            _target = target;
            _parseResult = parseResult;
            _wrapperFactory = wrapperFactory;
            _messageBox = messageBox;
        }

        public override void Fix()
        {
            var vbe = Selection.QualifiedName.Project.VBE;

            var refactoring = new MoveCloserToUsageRefactoring(_parseResult,
                new ActiveCodePaneEditor(vbe, _wrapperFactory), _messageBox);

            refactoring.Refactor(_target);
        }

        public override bool CanFixInModule { get { return false; } }
        public override bool CanFixInProject { get { return false; } }
    }
}
