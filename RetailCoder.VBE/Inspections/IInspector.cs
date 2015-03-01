using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using System.Threading.Tasks;

namespace Rubberduck.Inspections
{
    public interface IInspector
    {
        Task<IList<ICodeInspectionResult>> FindIssues(VBProject project);
        event EventHandler<InspectorIssuesFoundEventArg> IssuesFound;
    }

    public class InspectorIssuesFoundEventArg : EventArgs
    {
        private readonly int _count;
        public int Count { get { return _count; } }

        public InspectorIssuesFoundEventArg(int count)
        {
            _count = count;
        }
    }
}
