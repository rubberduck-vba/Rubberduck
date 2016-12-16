using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class WriteOnlyPropertyInspectionResult : InspectionResultBase
    {
        public WriteOnlyPropertyInspectionResult(IInspection inspection, Declaration target) 
            : base(inspection, target)
        {
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.WriteOnlyPropertyInspectionResultFormat, Target.IdentifierName).Captialize(); }
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return new QuickFixBase[]
                {
                    new WriteOnlyPropertyQuickFix(Context, Target),
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName)
                };
            }
        }
    }
}