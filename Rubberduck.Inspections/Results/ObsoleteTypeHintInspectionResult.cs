using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class ObsoleteTypeHintInspectionResult : InspectionResultBase
    {
        private readonly string _result;

        public ObsoleteTypeHintInspectionResult(IInspection inspection, string result, QualifiedContext qualifiedContext, Declaration declaration)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context, declaration)
        {
            _result = result;
        }

        public override string Description
        {
            get { return _result.Capitalize(); }
        }
    }
}
