using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Refactorings;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Results
{
    public class AssignedByValParameterInspectionResult : InspectionResultBase
    {
        private IEnumerable<QuickFixBase> _quickFixes;
        private readonly RubberduckParserState _parserState;
        private readonly IAssignedByValParameterQuickFixDialogFactory _dialogFactory;

        public AssignedByValParameterInspectionResult(IInspection inspection, Declaration target, RubberduckParserState parserState, IAssignedByValParameterQuickFixDialogFactory dialogFactory) 
            : base(inspection, target)
        {
            _dialogFactory = dialogFactory;
            _parserState = parserState;
        }

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
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new AssignedByValParameterMakeLocalCopyQuickFix(Target, QualifiedSelection, _parserState, _dialogFactory),
                    new PassParameterByReferenceQuickFix(Target, QualifiedSelection),
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }
    }
}
