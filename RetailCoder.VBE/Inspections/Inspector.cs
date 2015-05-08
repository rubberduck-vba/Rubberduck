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

        private static bool _isInspecting;

        public Inspector(IRubberduckParser parser, IEnumerable<IInspection> inspections)
        {
            _parser = parser;
            _parser.ParseStarted += _parser_ParseStarted;
            _parser.ParseCompleted += _parser_ParseCompleted;

            _inspections = inspections.ToList();
        }

        private void _parser_ParseCompleted(object sender, ParseCompletedEventArgs e)
        {
            if (!_isInspecting)
            {
                OnParseCompleted(e);
            }
        }

        private void _parser_ParseStarted(object sender, ParseStartedEventArgs e)
        {
            if (!_isInspecting)
            {
                OnParsing();
            }
        }

        public async Task<IList<ICodeInspectionResult>> FindIssuesAsync(VBProjectParseResult project)
        {
            _isInspecting = true;
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

            _isInspecting = false;
            return allIssues.ToList();
        }

        public void Parse(VBE vbe)
        {
            Task.Run(() => _parser.Parse(vbe));
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
        private void OnParsing()
        {
            var handler = Parsing;
            if (handler == null)
            {
                return;
            }

            handler(this, EventArgs.Empty);
        }

        public event EventHandler<ParseCompletedEventArgs> ParseCompleted;
        private void OnParseCompleted(ParseCompletedEventArgs args)
        {
            var handler = ParseCompleted;
            if (handler == null)
            {
                return;
            }

            handler(this, args);
        }
    }
}
