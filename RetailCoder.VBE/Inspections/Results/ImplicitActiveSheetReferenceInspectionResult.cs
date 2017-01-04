using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class ImplicitActiveSheetReferenceInspectionResult : InspectionResultBase
    {
        private readonly IdentifierReference _reference;
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public ImplicitActiveSheetReferenceInspectionResult(IInspection inspection, IdentifierReference reference)
            : base(inspection, reference.QualifiedModuleName, reference.Context)
        {
            _reference = reference;
            _quickFixes = new QuickFixBase[]
            {
                new IgnoreOnceQuickFix(reference.Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ImplicitActiveSheetReferenceInspectionResultFormat, _reference.Declaration.IdentifierName).Captialize(); }
        }
    }
}