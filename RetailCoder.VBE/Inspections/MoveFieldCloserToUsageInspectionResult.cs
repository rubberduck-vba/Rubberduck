using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class MoveFieldCloserToUsageInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public MoveFieldCloserToUsageInspectionResult(IInspection inspection, Declaration target, RubberduckParserState state, IMessageBox messageBox)
            : base(inspection, target)
        {
            _quickFixes = new[]
            {
                new MoveFieldCloserToUsageQuickFix(target.Context, target.QualifiedSelection, target, state, messageBox),
            };
        }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.MoveFieldCloserToUsageInspectionResultFormat, Target.IdentifierName);
            }
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    /// <summary>
    /// A code inspection quickfix that encapsulates a public field with a property
    /// </summary>
    public class MoveFieldCloserToUsageQuickFix : CodeInspectionQuickFix
    {
        private readonly Declaration _target;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public MoveFieldCloserToUsageQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target, RubberduckParserState state, IMessageBox messageBox)
            : base(context, selection, string.Format(InspectionsUI.MoveFieldCloserToUsageInspectionResultFormat, target.IdentifierName))
        {
            _target = target;
            _state = state;
            _messageBox = messageBox;
        }

        public override void Fix()
        {
            var vbe = Selection.QualifiedName.Project.VBE;

            var refactoring = new MoveCloserToUsageRefactoring(vbe, _state, _messageBox);

            refactoring.Refactor(_target);
        }

        public override bool CanFixInModule { get { return false; } }
        public override bool CanFixInProject { get { return false; } }
    }
}
