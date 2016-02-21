using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections
{
    public class NonReturningFunctionInspectionResult : InspectionResultBase
    {
        private readonly string _identifierName;
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public NonReturningFunctionInspectionResult(IInspection inspection,
            QualifiedContext<ParserRuleContext> qualifiedContext, 
            bool isInterfaceImplementation,
            string identifierName)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _identifierName = identifierName;
            _quickFixes = isInterfaceImplementation 
                ? new CodeInspectionQuickFix[] { }
                : new CodeInspectionQuickFix[]
                {
                    new ConvertToProcedureQuickFix(Context, QualifiedSelection),
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
                };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.NonReturningFunctionInspectionResultFormat, _identifierName);
            }
        }
    }
}