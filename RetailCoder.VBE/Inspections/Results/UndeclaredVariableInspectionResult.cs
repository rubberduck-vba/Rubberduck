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
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public UndeclaredVariableInspectionResult(IInspection inspection, Declaration target)
            : base(inspection, target)
        {
            _quickFixes = new QuickFixBase[]
            {
                new IntroduceLocalVariableQuickFix(target), 
                new IgnoreOnceQuickFix(target.Context, target.QualifiedSelection, inspection.AnnotationName), 
            };
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get { return _quickFixes; }
        }

        public override string Description { get { return string.Format(InspectionsUI.UndeclaredVariableInspectionResultFormat, Target.IdentifierName).Captialize(); } }
    }
}