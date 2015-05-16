using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections
{
    public class Inspector : IInspector
    {
        private readonly IRubberduckParser _parser;
        private readonly IList<IInspection> _inspections;

        public Inspector(IRubberduckParser parser, IEnumerable<IInspection> inspections)
        {
            _parser = parser;
            _parser.ParseStarted += _parser_ParseStarted;
            _parser.ParseCompleted += _parser_ParseCompleted;

            _inspections = inspections.ToList();
        }

        private void _parser_ParseCompleted(object sender, ParseCompletedEventArgs e)
        {
            OnParseCompleted(sender, e);
        }

        private void _parser_ParseStarted(object sender, ParseStartedEventArgs e)
        {
            OnParsing(sender);
        }

        public async Task<IList<ICodeInspectionResult>> FindIssuesAsync(VBProjectParseResult project)
        {
            await Task.Yield();

            OnReset();

            var allIssues = new ConcurrentBag<ICodeInspectionResult>();

            var inspections = _inspections.Where(inspection => inspection.Severity != CodeInspectionSeverity.DoNotShow)
                .Select(inspection =>
                    new Task(() =>
                    {
                        var inspectionResults = inspection.GetInspectionResults(project);
                        var results = inspectionResults as IList<CodeInspectionResultBase> ?? inspectionResults.ToList();

                        if (results.Any())
                        {
                            OnIssuesFound(results);

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

        public void Parse(VBE vbe, object owner)
        {
            Task.Run(() => _parser.Parse(vbe, owner));
        }

        public async Task<VBProjectParseResult> Parse(VBProject project, object owner)
        {
            return await Task.Run(() => _parser.Parse(project, owner));
        }

        public event EventHandler<InspectorIssuesFoundEventArg> IssuesFound;
        private void OnIssuesFound(IList<CodeInspectionResultBase> issues)
        {
            var handler = IssuesFound;
            if (handler == null)
            {
                return;
            }

            var args = new InspectorIssuesFoundEventArg(issues);
            handler(this, args);
        }

        public event EventHandler Reset;
        private void OnReset()
        {
            var handler = Reset;
            if (handler == null)
            {
                return;
            }

            handler(this, EventArgs.Empty);
        }

        public event EventHandler Parsing;
        private void OnParsing(object owner)
        {
            var handler = Parsing;
            if (handler == null)
            {
                return;
            }

            handler(owner, EventArgs.Empty);
        }

        public event EventHandler<ParseCompletedEventArgs> ParseCompleted;
        private void OnParseCompleted(object owner, ParseCompletedEventArgs args)
        {
            var handler = ParseCompleted;
            if (handler == null)
            {
                return;
            }

            handler(owner, args);
        }
    }
}
