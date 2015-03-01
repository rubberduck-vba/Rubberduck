using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Inspections
{
    public interface IInspector
    {
        IList<ICodeInspectionResult> FindIssues(VBProject project);
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
