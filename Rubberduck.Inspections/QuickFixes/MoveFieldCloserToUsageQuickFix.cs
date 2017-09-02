using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.UI;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class MoveFieldCloserToUsageQuickFix : QuickFixBase, IQuickFix
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public MoveFieldCloserToUsageQuickFix(RubberduckParserState state, InspectionLocator inspectionLocator, IMessageBox messageBox)
        {
            _state = state;
            _messageBox = messageBox;
            RegisterInspections(inspectionLocator.GetInspection<MoveFieldCloserToUsageInspection>());
        }

        public void Fix(IInspectionResult result)
        {
            var vbe = result.Target.Project.VBE;

            var refactoring = new MoveCloserToUsageRefactoring(vbe, _state, _messageBox);
            refactoring.Refactor(result.Target);
        }

        public string Description(IInspectionResult result)
        {
            return string.Format(InspectionsUI.MoveFieldCloserToUsageInspectionResultFormat, result.Target.IdentifierName);
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}