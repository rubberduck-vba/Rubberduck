using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections.Results
{
    public sealed class ImplicitVariantReturnTypeInspectionResult : InspectionResultBase
    {
        public ImplicitVariantReturnTypeInspectionResult(IInspection inspection, QualifiedContext<ParserRuleContext> qualifiedContext, Declaration target)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context, target) {}

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.ImplicitVariantReturnTypeInspectionResultFormat, Target.IdentifierName);
            }
        }

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
        }
    }
}
