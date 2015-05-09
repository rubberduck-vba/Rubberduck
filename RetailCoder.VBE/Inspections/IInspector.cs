using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections
{
    public interface IInspector
    {
        Task<IList<ICodeInspectionResult>> FindIssuesAsync(VBProjectParseResult project);
        void Parse(VBE vbe, object owner);
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
