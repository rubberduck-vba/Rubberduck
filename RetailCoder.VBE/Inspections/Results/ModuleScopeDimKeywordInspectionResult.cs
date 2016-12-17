using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class ModuleScopeDimKeywordInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public ModuleScopeDimKeywordInspectionResult(IInspection inspection, Declaration target) 
            : base(inspection, target)
        {
            _quickFixes = new QuickFixBase[]
            {
                new ChangeDimToPrivateQuickFix(target.Context, target.QualifiedSelection),
                //new IgnoreOnceQuickFix(Target.ParentDeclaration.Context, Target.ParentDeclaration.QualifiedSelection, Inspection.AnnotationName),
            };
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get { return _quickFixes; }
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ModuleScopeDimKeywordInspectionResultFormat, Target.IdentifierName).Captialize(); }
        }
    }
}