using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class EncapsulateFieldQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;
        private readonly IVBE _vbe;
        private readonly IRewritingManager _rewritingManager;
        private readonly IIndenter _indenter;
        private readonly IRefactoringPresenterFactory _factory;
        
        public EncapsulateFieldQuickFix(RubberduckParserState state, IVBE vbe, IIndenter indenter, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager)
            : base(typeof(EncapsulatePublicFieldInspection))
        {
            _state = state;
            _vbe = vbe;
            _rewritingManager = rewritingManager;
            _indenter = indenter;
            _factory = factory;
        }

        //The rewriteSession is optional since it is not used in this particular quickfix because it is a refactoring quickfix.
        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession = null)
        {
            var refactoring = new EncapsulateFieldRefactoring(_state, _vbe, _indenter, _factory, _rewritingManager);
            refactoring.Refactor(result.Target);
        }

        public override string Description(IInspectionResult result)
        {
            return string.Format(Resources.Inspections.QuickFixes.EncapsulatePublicFieldInspectionQuickFix, result.Target.IdentifierName);
        }

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
    }
}