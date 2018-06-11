using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
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

        public RemoveUnusedParameterQuickFix(IVBE vbe, RubberduckParserState state, IRefactoringPresenterFactory factory)
            : base(typeof(ParameterNotUsedInspection))
        {
            _vbe = vbe;
            _state = state;
            _factory = factory;
        }

        public override void Fix(IInspectionResult result)
        {
            var refactoring = new RemoveParametersRefactoring(_state, _vbe, _factory);
            refactoring.QuickFix(_state, result.QualifiedSelection);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveUnusedParameterQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}