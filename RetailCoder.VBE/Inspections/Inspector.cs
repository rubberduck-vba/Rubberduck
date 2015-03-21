using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.VBA;

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

        public async Task<IList<ICodeInspectionResult>> FindIssuesAsync(VBProject project)
        {
            await Task.Yield();

            RaiseResetEvent();

            var code = new VBProjectParseResult(_parser.Parse(project));
            var allIssues = new ConcurrentBag<ICodeInspectionResult>();

            var inspections = _inspections.Where(inspection => inspection.Severity != CodeInspectionSeverity.DoNotShow)
                .Select(inspection =>
                    new Task(() =>
                    {
                        var inspectionResults = inspection.GetInspectionResults(code);
                        var results = inspectionResults as IList<CodeInspectionResultBase> ?? inspectionResults.ToList();

                        if (results.Any())
                        {
                            RaiseIssuesFoundEvent(results);

                            foreach (var inspectionResult in results)
                            {
                                allIssues.Add(inspectionResult);
                            }
                        }
                    })).ToArray();

            foreach (var inspection in inspections)
            {
                inspection.Start();
            }

            Task.WaitAll(inspections);

            return allIssues.ToList();
        }

        public event EventHandler<InspectorIssuesFoundEventArg> IssuesFound;
        private void RaiseIssuesFoundEvent(IList<CodeInspectionResultBase> issues)
        {
            var handler = IssuesFound;
            if (handler == null)
            {
                return;
            }

            var args = new InspectorIssuesFoundEventArg(issues);
            handler(this, args);
        }

        public event EventHandler<EventArgs> Reset;
        private void RaiseResetEvent()
        {
            var handler = Reset;
            if (handler == null)
            {
                return;
            }

            handler(this, EventArgs.Empty);
        }
    }
}
