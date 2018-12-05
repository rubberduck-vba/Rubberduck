using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveUnusedParameterQuickFix : QuickFixBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IRefactoringPresenterFactory _factory;
        private readonly IRewritingManager _rewritingManager;

        public RemoveUnusedParameterQuickFix(IVBE vbe, RubberduckParserState state, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager)
            : base(typeof(ParameterNotUsedInspection))
        {
            _vbe = vbe;
            _state = state;
            _factory = factory;
            _rewritingManager = rewritingManager;
        }

        //The rewriteSession is optional since it is not used in this particular quickfix because it is a refactoring quickfix.
        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession = null)
        {
            var refactoring = new RemoveParametersRefactoring(_state, _vbe, _factory, _rewritingManager);
            refactoring.QuickFix(_state, result.QualifiedSelection);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveUnusedParameterQuickFix;

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
    }
}