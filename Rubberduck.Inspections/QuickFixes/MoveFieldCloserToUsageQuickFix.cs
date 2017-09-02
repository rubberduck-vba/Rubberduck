using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.UI;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class MoveFieldCloserToUsageQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public MoveFieldCloserToUsageQuickFix(RubberduckParserState state, IMessageBox messageBox)
            : base(typeof(MoveFieldCloserToUsageInspection))
        {
            _state = state;
            _messageBox = messageBox;
        }

        public override void Fix(IInspectionResult result)
        {
            var vbe = result.Target.Project.VBE;

            var refactoring = new MoveCloserToUsageRefactoring(vbe, _state, _messageBox);
            refactoring.Refactor(result.Target);
        }

        public override string Description(IInspectionResult result)
        {
            return string.Format(InspectionsUI.MoveFieldCloserToUsageInspectionResultFormat, result.Target.IdentifierName);
        }

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}