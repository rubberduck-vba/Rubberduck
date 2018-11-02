using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Refactorings.EncapsulateField;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class EncapsulateFieldQuickFix : QuickFixBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;
        private readonly IIndenter _indenter;

        public EncapsulateFieldQuickFix(IVBE vbe, RubberduckParserState state, IIndenter indenter, IRewritingManager rewritingManager)
            : base(typeof(EncapsulatePublicFieldInspection))
        {
            _vbe = vbe;
            _state = state;
            _rewritingManager = rewritingManager;
            _indenter = indenter;
        }

        //The rewriteSession is optional since it is not used in this particular quickfix because it is a refactoring quickfix.
        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession = null)
        {
            using (var view = new EncapsulateFieldDialog(new EncapsulateFieldViewModel(_state, _indenter)))
            {
                var factory = new EncapsulateFieldPresenterFactory(_vbe, _state, view);
                var refactoring = new EncapsulateFieldRefactoring(_vbe, _indenter, factory, _rewritingManager);
                refactoring.Refactor(result.Target);
            }
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