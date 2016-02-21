using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections
{
    public class NonReturningFunctionInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public NonReturningFunctionInspectionResult(IInspection inspection, QualifiedContext<ParserRuleContext> qualifiedContext, 
            bool isInterfaceImplementation)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context)
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

        public override string Description
        {
            get
            {
                // bug NullReferenceException thrown here - null Target
                return string.Format(InspectionsUI.NonReturningFunctionInspectionResultFormat, Target.IdentifierName);
            }
        }
    }
}