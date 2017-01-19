using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// A code inspection quickfix that encapsulates a public field with a property
    /// </summary>
    public class MoveFieldCloserToUsageQuickFix : QuickFixBase
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
            var vbe = _target.Project.VBE;

            var refactoring = new MoveCloserToUsageRefactoring(vbe, _state, _messageBox);

            refactoring.Refactor(_target);
        }

        public override bool CanFixInModule { get { return false; } }
        public override bool CanFixInProject { get { return false; } }
    }
}