using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections.Results
{
    public class ImplicitPublicMemberInspectionResult : InspectionResultBase
    {
        public ImplicitPublicMemberInspectionResult(IInspection inspection, QualifiedContext<ParserRuleContext> qualifiedContext, Declaration item)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context, item) {}

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.ImplicitPublicMemberInspectionResultFormat, Target.IdentifierName);
            }
        }

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
        }
    }
}
