using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

namespace Rubberduck.Inspections.Results
{
    public class ObsoleteCommentSyntaxInspectionResult : InspectionResultBase
    {
        private IEnumerable<IQuickFix> _quickFixes;

        public ObsoleteCommentSyntaxInspectionResult(IInspection inspection, QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context)
        { }

        public override IEnumerable<IQuickFix> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new IQuickFix[]
                {
                    new ReplaceObsoleteCommentMarkerQuickFix(Context, QualifiedSelection),
                    new RemoveCommentQuickFix(Context, QualifiedSelection), 
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }

        public override string Description
        {
            get { return InspectionsUI.ObsoleteCommentSyntaxInspectionResultFormat; }
        }
    }
}
