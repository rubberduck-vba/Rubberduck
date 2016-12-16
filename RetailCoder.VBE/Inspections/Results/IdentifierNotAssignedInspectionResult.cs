using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class IdentifierNotAssignedInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public IdentifierNotAssignedInspectionResult(IInspection inspection, Declaration target,
            ParserRuleContext context)
            : base(inspection, target)
        {
            _quickFixes = new QuickFixBase[]
            {
                new RemoveUnassignedIdentifierQuickFix(Context, QualifiedSelection, target), 
                new IgnoreOnceQuickFix(context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.VariableNotAssignedInspectionResultFormat, Target.IdentifierName).Captialize(); }
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }
    }
}
