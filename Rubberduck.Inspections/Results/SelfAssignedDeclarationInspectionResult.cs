using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class SelfAssignedDeclarationInspectionResult : InspectionResultBase
    {
        public SelfAssignedDeclarationInspectionResult(IInspection inspection, Declaration target)
            : base(inspection, target)
        {
        }

        public override IEnumerable<IQuickFix> QuickFixes
        {
            get
            {
                return new IQuickFix[]
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
