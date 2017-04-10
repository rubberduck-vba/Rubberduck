using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections.Results
{
    public sealed class NonReturningFunctionInspectionResult : InspectionResultBase
    {
        public NonReturningFunctionInspectionResult(IInspection inspection, QualifiedContext<ParserRuleContext> qualifiedContext, Declaration target)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context, target) {}

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.NonReturningFunctionInspectionResultFormat, Target.IdentifierName).Capitalize();
            }
        }

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
        }
    }
}
