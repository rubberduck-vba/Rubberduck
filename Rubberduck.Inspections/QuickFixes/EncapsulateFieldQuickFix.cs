using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Refactorings.EncapsulateField;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class EncapsulateFieldQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;
        private readonly IIndenter _indenter;

        public EncapsulateFieldQuickFix(RubberduckParserState state, IIndenter indenter)
            : base(typeof(EncapsulatePublicFieldInspection))
        {
            _state = state;
            _indenter = indenter;
        }

        public override void Fix(IInspectionResult result)
        {
            var vbe = result.Target.Project.VBE;

            using (var view = new EncapsulateFieldDialog(new EncapsulateFieldViewModel(_state, _indenter)))
            {
                var factory = new EncapsulateFieldPresenterFactory(vbe, _state, view);
                var refactoring = new EncapsulateFieldRefactoring(vbe, _indenter, factory);
                refactoring.Refactor(result.Target);
            }
        }

        public override string Description(IInspectionResult result)
        {
            return string.Format(InspectionsUI.EncapsulatePublicFieldInspectionQuickFix, result.Target.IdentifierName);
        }

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
    }
}