using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections.Results
{
    public class DefaultProjectNameInspectionResult : InspectionResultBase
    {
        public DefaultProjectNameInspectionResult(IInspection inspection, Declaration target, RubberduckParserState state, IMessageBox messageBox)
            : base(inspection, target) {}

        public override string Description
        {
            get { return Inspection.Description.Capitalize(); }
        }
    }
}
