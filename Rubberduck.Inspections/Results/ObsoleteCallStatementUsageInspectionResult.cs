using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

namespace Rubberduck.Inspections.Results
{
    public class ObsoleteCallStatementUsageInspectionResult : InspectionResultBase
    {
        public ObsoleteCallStatementUsageInspectionResult(IInspection inspection, QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context) {}

        public override string Description
        {
            get { return InspectionsUI.ObsoleteCallStatementInspectionResultFormat.Capitalize(); }
        }
    }
}
