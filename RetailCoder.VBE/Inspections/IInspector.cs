using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections
{
    public interface IInspector
    {
        Task<IList<ICodeInspectionResult>> FindIssuesAsync(VBProjectParseResult project, CancellationToken token);
        void Parse(VBE vbe, object owner);
        Task<VBProjectParseResult> Parse(VBProject project, object owner);
        event EventHandler<InspectorIssuesFoundEventArg> IssuesFound;
        event EventHandler Reset;
        event EventHandler Parsing;
        event EventHandler<ParseCompletedEventArgs> ParseCompleted;
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
