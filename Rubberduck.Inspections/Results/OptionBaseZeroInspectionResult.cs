using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

namespace Rubberduck.Inspections.Results
{
    public class OptionBaseZeroInspectionResult : InspectionResultBase
    {
        public OptionBaseZeroInspectionResult(IInspection inspection, QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context) {}

        public override string Description
        {
            get { return string.Format(InspectionsUI.OptionBaseZeroInspectionResultFormat.Capitalize(), QualifiedName.ComponentName); }
        }
    }
}
