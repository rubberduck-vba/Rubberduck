using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class UndeclaredVariableInspectionResult : InspectionResultBase
    {
        private IEnumerable<QuickFixBase> _quickFixes;

        public UndeclaredVariableInspectionResult(IInspection inspection, Declaration target)
            : base(inspection, target)
        { }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new IntroduceLocalVariableQuickFix(Target), 
                    new IgnoreOnceQuickFix(Target.Context, Target.QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }

        public override string Description { get { return string.Format(InspectionsUI.UndeclaredVariableInspectionResultFormat, Target.IdentifierName).Captialize(); } }
    }
}