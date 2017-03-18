using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

namespace Rubberduck.Inspections.Results
{
    public class EmptyStringLiteralInspectionResult : InspectionResultBase
    {
        private IEnumerable<IQuickFix> _quickFixes;

        public EmptyStringLiteralInspectionResult(IInspection inspection, QualifiedContext qualifiedContext)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context)
        { }

        public override IEnumerable<IQuickFix> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new IQuickFix[]
                {
                    new RepaceEmptyStringLiteralStatementQuickFix(Context, QualifiedSelection),
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }

        public override string Description
        {
            get { return InspectionsUI.EmptyStringLiteralInspectionResultFormat.Captialize(); }
        }
    }
}
