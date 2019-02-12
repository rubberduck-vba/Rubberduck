using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveUnusedParameterQuickFix : QuickFixBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IRefactoringPresenterFactory _factory;
        private readonly IRewritingManager _rewritingManager;
        private readonly ISelectionService _selectionService;

        public RemoveUnusedParameterQuickFix(IDeclarationFinderProvider declarationFinderProvider, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService)
            : base(typeof(ParameterNotUsedInspection))
        {
            _declarationFinderProvider = declarationFinderProvider;
            _factory = factory;
            _rewritingManager = rewritingManager;
            _selectionService = selectionService;
        }

        //The rewriteSession is optional since it is not used in this particular quickfix because it is a refactoring quickfix.
        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession = null)
        {
            var refactoring = new RemoveParametersRefactoring(_declarationFinderProvider, _factory, _rewritingManager, _selectionService);
            refactoring.QuickFix(result.QualifiedSelection);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveUnusedParameterQuickFix;

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
    }
}