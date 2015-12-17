using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public interface IInspector
    {
        Task<IList<ICodeInspectionResult>> FindIssuesAsync(IRubberduckParserState state, CancellationToken token);
    }

    public class InspectorIssuesFoundEventArg : EventArgs
    {
        private readonly IList<CodeInspectionResultBase> _issues;
        public IList<CodeInspectionResultBase> Issues { get { return _issues; } }

        public InspectorIssuesFoundEventArg(IList<CodeInspectionResultBase> issues)
        {
            _issues = issues;
        }
    }
}
