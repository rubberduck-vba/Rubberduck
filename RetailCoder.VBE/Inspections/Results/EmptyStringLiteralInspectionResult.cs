using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections.Results
{
    public class EmptyStringLiteralInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public EmptyStringLiteralInspectionResult(IInspection inspection, QualifiedContext<VBAParser.LiteralExpressionContext> qualifiedContext)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _quickFixes = new QuickFixBase[]
            {
                new RepaceEmptyStringLiteralStatementQuickFix(Context, QualifiedSelection),
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName)
            };
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return InspectionsUI.EmptyStringLiteralInspectionResultFormat.Captialize(); }
        }
    }
}
