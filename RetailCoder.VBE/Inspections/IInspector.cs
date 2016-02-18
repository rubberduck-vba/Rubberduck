using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public interface IInspector
    {
        Task<IList<ICodeInspectionResult>> FindIssuesAsync(RubberduckParserState state, CancellationToken token);
    }

    public class InspectorIssuesFoundEventArg : EventArgs
    {
        private readonly IList<InspectionResultBase> _issues;
        public IList<InspectionResultBase> Issues { get { return _issues; } }

        public InspectorIssuesFoundEventArg(IList<InspectionResultBase> issues)
        {
            _issues = issues;
        }
    }
}
