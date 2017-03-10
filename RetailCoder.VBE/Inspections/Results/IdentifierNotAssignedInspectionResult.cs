using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class IdentifierNotAssignedInspectionResult : InspectionResultBase
    {
        private IEnumerable<IQuickFix> _quickFixes;
        private readonly ParserRuleContext _context;
        private readonly TokenStreamRewriter _rewriter;

        public IdentifierNotAssignedInspectionResult(IInspection inspection, Declaration target, ParserRuleContext context, TokenStreamRewriter rewriter)
            : base(inspection, target)
        {
            _context = context;
            _rewriter = rewriter;
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.VariableNotAssignedInspectionResultFormat, Target.IdentifierName).Captialize(); }
        }

        public override IEnumerable<IQuickFix> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new IQuickFix[]
                {
                    new RemoveUnassignedIdentifierQuickFix(Context, QualifiedSelection, Target, _rewriter), 
                    new IgnoreOnceQuickFix(_context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }
    }
}
