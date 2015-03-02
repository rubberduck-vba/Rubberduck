using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using System.Threading.Tasks;

namespace Rubberduck.Inspections
{
    public interface IInspector
    {
        Task<IList<ICodeInspectionResult>> FindIssuesAsync(VBProject project);
        event EventHandler<InspectorIssuesFoundEventArg> IssuesFound;
        event EventHandler<EventArgs> Reset;
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
