using Rubberduck.Common;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Inspections.Abstract;

namespace Rubberduck.Formatters
{
    public class InspectionResultFormatter : IExportable
    {
        private readonly IInspectionResult _inspectionResult;

        public InspectionResultFormatter(IInspectionResult inspectionResult)
        {
            _inspectionResult = inspectionResult;
        }

        public object[] ToArray()
        {
            return ((InspectionResultBase)_inspectionResult).ToArray();
        }

        public string ToClipboardString()
        {
            return ((InspectionResultBase)_inspectionResult).ToClipboardString();
        }
    }
}
