using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

namespace Rubberduck.Inspections.Results
{
    public class OptionBaseInspectionResult : InspectionResultBase
    {
        public OptionBaseInspectionResult(IInspection inspection, QualifiedContext<ParserRuleContext> context)
            : base(inspection, context.ModuleName, context.Context) {}

        public override string Description
        {
            get { return string.Format(InspectionsUI.OptionBaseInspectionResultFormat, QualifiedName.ComponentName); }
        }
    }
}
