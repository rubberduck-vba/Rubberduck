using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class EncapsulateFieldQuickFix : QuickFixBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IIndenter _indenter;
        private readonly IRefactoringPresenterFactory<IEncapsulateFieldPresenter> _factory;

        public EncapsulateFieldQuickFix(IVBE vbe, RubberduckParserState state, IIndenter indenter, IRefactoringPresenterFactory<IEncapsulateFieldPresenter> factory)
            : base(typeof(EncapsulatePublicFieldInspection))
        {
            _vbe = vbe;
            _state = state;
            _indenter = indenter;
            _factory = factory;
        }

        public override void Fix(IInspectionResult result)
        {
            var refactoring = new EncapsulateFieldRefactoring(_vbe, _indenter, _factory);
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