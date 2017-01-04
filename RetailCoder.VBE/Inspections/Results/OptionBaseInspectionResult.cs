using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class OptionBaseInspectionResult : InspectionResultBase
    {
        public OptionBaseInspectionResult(IInspection inspection, Declaration target)
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
            get { return string.Format(InspectionsUI.OptionBaseInspectionResultFormat, QualifiedName.ComponentName); }
        }
    }
}
