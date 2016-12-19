using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections.Results
{
    public sealed class NonReturningFunctionInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public NonReturningFunctionInspectionResult(IInspection inspection,
            QualifiedContext<ParserRuleContext> qualifiedContext, 
            Declaration target, bool canConvertToProcedure)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context, target)
        {
            _quickFixes = canConvertToProcedure
                ? new QuickFixBase[] 
                {
                    new ConvertToProcedureQuickFix(Context, QualifiedSelection, target),
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName),
                }
                : new QuickFixBase[]
                {
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
                };
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

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
