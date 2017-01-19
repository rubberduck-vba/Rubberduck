using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    public class MultipleDeclarationsInspectionResult : InspectionResultBase
    {
        private IEnumerable<QuickFixBase> _quickFixes;
        private readonly QualifiedContext<ParserRuleContext> _qualifiedContext;

        public MultipleDeclarationsInspectionResult(IInspection inspection, QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _qualifiedContext = qualifiedContext;
        }

        public override string Description
        {
            get { return InspectionsUI.MultipleDeclarationsInspectionResultFormat.Captialize(); }
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new SplitMultipleDeclarationsQuickFix(Context, QualifiedSelection), 
                    new IgnoreOnceQuickFix(_qualifiedContext.Context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }

        private new QualifiedSelection QualifiedSelection
        {
            get
            {
                ParserRuleContext context;
                if (Context is VBAParser.ConstStmtContext)
                {
                    context = Context;
                }
                else
                {
                    context = Context.Parent as ParserRuleContext;
                }
                var selection = context.GetSelection();
                return new QualifiedSelection(QualifiedName, selection);
            }
        }
    }
}
