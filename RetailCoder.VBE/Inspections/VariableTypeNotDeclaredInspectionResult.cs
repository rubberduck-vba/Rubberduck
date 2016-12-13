using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class VariableTypeNotDeclaredInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public VariableTypeNotDeclaredInspectionResult(IInspection inspection, Declaration target)
            : base(inspection, target)
        {
            _quickFixes = new CodeInspectionQuickFix[]
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

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get {return _quickFixes; } }
    }
}
