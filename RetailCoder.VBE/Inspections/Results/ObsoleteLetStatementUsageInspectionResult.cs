using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections.Results
{
    public class ObsoleteLetStatementUsageInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public ObsoleteLetStatementUsageInspectionResult(IInspection inspection, QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _quickFixes = new QuickFixBase[]
            {
                new RemoveExplicitLetStatementQuickFix(Context, QualifiedSelection), 
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override string Description
        {
            get { return InspectionsUI.ObsoleteLetStatementInspectionResultFormat; }
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get {return _quickFixes; } }
    }
}
