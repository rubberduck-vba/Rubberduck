using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections
{
    public class NonReturningFunctionInspectionResult : CodeInspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public NonReturningFunctionInspectionResult(IInspection inspection, string result, QualifiedContext<ParserRuleContext> qualifiedContext, 
            bool isInterfaceImplementation)
            : base(inspection, result, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _quickFixes = isInterfaceImplementation 
                ? new CodeInspectionQuickFix[] { }
                : new CodeInspectionQuickFix[]
                {
                    new ConvertToProcedureQuickFix(Context, QualifiedSelection),
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
                };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }
}