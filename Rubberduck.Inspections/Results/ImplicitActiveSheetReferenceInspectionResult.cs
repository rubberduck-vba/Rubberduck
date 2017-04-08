using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class ImplicitActiveSheetReferenceInspectionResult : InspectionResultBase
    {
        private readonly IdentifierReference _reference;
        private IEnumerable<IQuickFix> _quickFixes;

        public ImplicitActiveSheetReferenceInspectionResult(IInspection inspection, IdentifierReference reference)
            : base(inspection, reference.QualifiedModuleName, reference.Context)
        {
            _reference = reference;
        }

        public override IEnumerable<IQuickFix> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new IQuickFix[]
                {
                    new IgnoreOnceQuickFix(_reference.Context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ImplicitActiveSheetReferenceInspectionResultFormat, _reference.Declaration.IdentifierName).Captialize(); }
        }
    }
}