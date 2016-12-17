using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.QuickFixes
{
    public class SelfAssignedDeclarationInspectionResult : InspectionResultBase
    {
        public SelfAssignedDeclarationInspectionResult(IInspection inspection, Declaration target)
            : base(inspection, target)
        {
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return new QuickFixBase[]
                {
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName)
                };
            }
        }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.SelfAssignedDeclarationInspectionResultFormat, Target.IdentifierName).Captialize();
            }
        }
    }
}
