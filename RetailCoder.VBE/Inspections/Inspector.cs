using Microsoft.Vbe.Interop;
using Rubberduck.VBA;
using Rubberduck.VBA.Nodes;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Rubberduck.Inspections
{
    public class Inspector : IInspector
    {
        private readonly IRubberduckParser _parser;
        private readonly IList<IInspection> _inspections;

        public Inspector(IRubberduckParser parser, IEnumerable<IInspection> inspections)
        {
            _parser = parser;
            _inspections = inspections.ToList();
        }

        public async Task<IList<ICodeInspectionResult>> FindIssues(VBProject project)
        {
            await Task.Yield();

            var code = new VBProjectParseResult(_parser.Parse(project));
            var results = new ConcurrentBag<ICodeInspectionResult>();

            var inspections = _inspections.Where(inspection => inspection.Severity != CodeInspectionSeverity.DoNotShow)
                .Select(inspection =>
                    new Task(() =>
                    {
                        var result = inspection.GetInspectionResults(code);
                        var count = result.Count();
                        if (count > 0)
                        {
                            RaiseIssuesFound(count);

                            foreach (var inspectionResult in result)
                            {
                                results.Add(inspectionResult);
                            }
                        }
                    })).ToArray();

            foreach (var inspection in inspections)
            {
                inspection.Start();
            }

            Task.WaitAll(inspections);

            return results.ToList();
        }

        public event EventHandler<InspectorIssuesFoundEventArg> IssuesFound;
        private void RaiseIssuesFound(int count)
        {
            var handler = IssuesFound;
            if (handler == null)
            {
                return;
            }

            var args = new InspectorIssuesFoundEventArg(count);
            handler(this, args);
        }
    }
}
