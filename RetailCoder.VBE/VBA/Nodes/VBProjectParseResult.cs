using System.Collections.Generic;
using Rubberduck.Inspections;

namespace Rubberduck.VBA.Nodes
{
    public class VBProjectParseResult
    {
        public VBProjectParseResult(IEnumerable<VBComponentParseResult> parseResults)
        {
            _parseResults = parseResults;
            _inspector = new IdentifierUsageInspector(_parseResults);
        }

        private readonly IEnumerable<VBComponentParseResult> _parseResults;
        public IEnumerable<VBComponentParseResult> ComponentParseResults { get { return _parseResults; } }

        private readonly IdentifierUsageInspector _inspector;
        public IdentifierUsageInspector IdentifierUsageInspector { get { return _inspector; } }
    }
}