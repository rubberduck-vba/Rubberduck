using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.Inspections.Results
{
    public class AssignedByValParameterInspectionResult : InspectionResultBase
    {
        private IEnumerable<QuickFixBase> _quickFixes;

        public AssignedByValParameterInspectionResult(IInspection inspection, Declaration target) : base(inspection, target) { }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.AssignedByValParameterInspectionResultFormat, Target.IdentifierName).Captialize();
            }
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                IAssignedByValParameterQuickFixDialogFactory factory = new AssignedByValParameterQuickFixDialogFactory();
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new AssignedByValParameterMakeLocalCopyQuickFix(Target, QualifiedSelection, factory),
                    new PassParameterByReferenceQuickFix(Target, QualifiedSelection),
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }
    }
}
