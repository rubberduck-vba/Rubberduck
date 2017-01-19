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
        private IEnumerable<QuickFixBase> _quickFixes;
        private readonly ParserRuleContext _context;

        public IdentifierNotAssignedInspectionResult(IInspection inspection, Declaration target, ParserRuleContext context) : base(inspection, target)
        {
            _context = context;
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.VariableNotAssignedInspectionResultFormat, Target.IdentifierName).Captialize(); }
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new RemoveUnassignedIdentifierQuickFix(Context, QualifiedSelection, Target), 
                    new IgnoreOnceQuickFix(_context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }
    }
}
