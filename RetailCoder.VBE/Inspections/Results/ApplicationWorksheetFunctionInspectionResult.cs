using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    public class ApplicationWorksheetFunctionInspectionResult : InspectionResultBase
    {
        private readonly QualifiedSelection _qualifiedSelection;
        private readonly string _memberName;
        private IEnumerable<QuickFixBase> _quickFixes;

        public ApplicationWorksheetFunctionInspectionResult(IInspection inspection, QualifiedSelection qualifiedSelection, string memberName)
            : base(inspection, qualifiedSelection.QualifiedName)
        {
            _memberName = memberName;
            _qualifiedSelection = qualifiedSelection;
        }

        public override QualifiedSelection QualifiedSelection
        {
            get { return _qualifiedSelection; }
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new IgnoreOnceQuickFix(null, _qualifiedSelection, Inspection.AnnotationName),
                    new ApplicationWorksheetFunctionQuickFix(_qualifiedSelection, _memberName)
                });
            }
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ApplicationWorksheetFunctionInspectionResultFormat, _memberName).Captialize(); }
        }
    }
}
