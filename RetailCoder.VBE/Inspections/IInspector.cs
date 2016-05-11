using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public interface IInspector
    {
        Task<IEnumerable<ICodeInspectionResult>> FindIssuesAsync(RubberduckParserState state, CancellationToken token);
    }

    public class InspectorIssuesFoundEventArg : EventArgs
    {
        private readonly IEnumerable<InspectionResultBase> _issues;
        public IEnumerable<InspectionResultBase> Issues { get { return _issues; } }

        public InspectorIssuesFoundEventArg(IEnumerable<InspectionResultBase> issues)
        {
            _issues = issues;
        }
    }
}
