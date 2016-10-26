using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class ModuleScopeDimKeywordInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ModuleScopeDimKeywordInspectionResult(IInspection inspection, Declaration target) 
            : base(inspection, target)
        {
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new ChangeDimToPrivateQuickFix(target.Context, target.QualifiedSelection),
                //new IgnoreOnceQuickFix(Target.ParentDeclaration.Context, Target.ParentDeclaration.QualifiedSelection, Inspection.AnnotationName),
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes
        {
            get { return _quickFixes; }
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ModuleScopeDimKeywordInspectionResultFormat, Target.IdentifierName); }
        }
    }
}