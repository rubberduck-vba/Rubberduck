using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public sealed class NonReturningFunctionInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public NonReturningFunctionInspectionResult(IInspection inspection,
            QualifiedContext<ParserRuleContext> qualifiedContext, 
            bool isInterfaceImplementation,
            Declaration target)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context, target)
        {
            _quickFixes = isInterfaceImplementation 
                ? new CodeInspectionQuickFix[] 
                {
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName),
                }
                : new CodeInspectionQuickFix[]
                {
                    new ConvertToProcedureQuickFix(Context, QualifiedSelection, target),
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
                };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.NonReturningFunctionInspectionResultFormat, Target.IdentifierName).Captialize();
            }
        }

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
        }
    }
}
