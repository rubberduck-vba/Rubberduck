using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class VariableTypeNotDeclaredInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public VariableTypeNotDeclaredInspectionResult(IInspection inspection, Declaration target)
            : base(inspection, target)
        {
            _quickFixes = new QuickFixBase[]
            {
                new DeclareAsExplicitVariantQuickFix(Context, QualifiedSelection), 
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.ImplicitVariantDeclarationInspectionResultFormat, 
                    Target.DeclarationType,
                    Target.IdentifierName).Captialize();
            }
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get {return _quickFixes; } }
    }
}
