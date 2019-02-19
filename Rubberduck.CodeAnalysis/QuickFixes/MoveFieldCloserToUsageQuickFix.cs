using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class MoveFieldCloserToUsageQuickFix : QuickFixBase
    {
        private readonly ISelectionService _selectionService;
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;
        private readonly IMessageBox _messageBox;

        public MoveFieldCloserToUsageQuickFix(RubberduckParserState state, IMessageBox messageBox, IRewritingManager rewritingManager, ISelectionService selectionService)
            : base(typeof(MoveFieldCloserToUsageInspection))
        {
            _selectionService = selectionService;
            _state = state;
            _rewritingManager = rewritingManager;
            _messageBox = messageBox;
        }

        //The rewriteSession is optional since it is not used in this particular quickfix because it is a refactoring quickfix.
        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession = null)
        {
            var refactoring = new MoveCloserToUsageRefactoring(_state, _messageBox, _rewritingManager, _selectionService);
            refactoring.Refactor(result.Target);
        }

        public override string Description(IInspectionResult result)
        {
            return string.Format(InspectionResults.MoveFieldCloserToUsageInspection, result.Target.IdentifierName);
        }

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
    }
}