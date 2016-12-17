using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class AssignedByValParameterInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public AssignedByValParameterInspectionResult(IInspection inspection, Declaration target)
            : base(inspection, target)
        {
            _quickFixes = new QuickFixBase[]
            {
                new PassParameterByReferenceQuickFix(target.Context, QualifiedSelection),
                new IgnoreOnceQuickFix(Context, QualifiedSelection, inspection.AnnotationName)
            };
        }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.AssignedByValParameterInspectionResultFormat, Target.IdentifierName).Captialize();
            }
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }
    }
}
