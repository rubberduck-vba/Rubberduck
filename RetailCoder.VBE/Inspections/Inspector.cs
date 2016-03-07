using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;

namespace Rubberduck.Inspections
{
    public class Inspector : IInspector, IDisposable
    {
        private readonly IGeneralConfigService _configService;
        private readonly IEnumerable<IInspection> _inspections;

        public Inspector(IGeneralConfigService configService, IEnumerable<IInspection> inspections)
        {
            _inspections = inspections;

            _configService = configService;
            configService.LanguageChanged += ConfigServiceLanguageChanged;
            UpdateInspectionSeverity();
        }

        private void ConfigServiceLanguageChanged(object sender, EventArgs e)
        {
            UpdateInspectionSeverity();
        }

        private void UpdateInspectionSeverity()
        {
            var config = _configService.LoadConfiguration();

            foreach (var inspection in _inspections)
            {
                foreach (var setting in config.UserSettings.CodeInspectionSettings.CodeInspections)
                {
                    if (inspection.Description == setting.Description)
                    {
                        inspection.Severity = setting.Severity;
                    }
                }
            }
        }

        public async Task<IList<ICodeInspectionResult>> FindIssuesAsync(RubberduckParserState state, CancellationToken token)
        {
            if (state == null || !state.AllUserDeclarations.Any())
            {
                return new ICodeInspectionResult[]{};
            }

            await Task.Yield();

            UpdateInspectionSeverity();
            //OnReset();

            var allIssues = new ConcurrentBag<ICodeInspectionResult>();

            var inspections = _inspections.Where(inspection => inspection.Severity != CodeInspectionSeverity.DoNotShow)
                .Select(inspection =>
                    new Task(() =>
                    {
                        token.ThrowIfCancellationRequested();
                        Console.WriteLine("Running inspection {0} in thread {1}", inspection.Name, Thread.CurrentThread.ManagedThreadId);
                        var inspectionResults = inspection.GetInspectionResults();
                        var results = inspectionResults as IList<InspectionResultBase> ?? inspectionResults.ToList();

                        if (results.Any())
                        {
                            //OnIssuesFound(results);
                            Console.WriteLine("Inspection '{0}' returned {1} results.", inspection.Name, results.Count);

                            foreach (var inspectionResult in results)
                            {
                                allIssues.Add(inspectionResult);
                            }
                        }
                    })).ToArray();

            try
            {
                var stopwatch = Stopwatch.StartNew();
                Console.WriteLine("Starting code inspections in thread {0}", Thread.CurrentThread.ManagedThreadId);
                foreach (var inspection in inspections)
                {
                    await inspection;
                }

                //Task.WaitAll(inspections);
                Console.WriteLine("Code inspections completed ({0}ms) in thread {1}", stopwatch.ElapsedMilliseconds, Thread.CurrentThread.ManagedThreadId);
            }
            catch (AggregateException exceptions)
            {
                Console.WriteLine(exceptions);
            }
            return allIssues.ToList();
        }

        public void Dispose()
        {
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposing) { return; }

            if (_configService != null)
            {
                _configService.LanguageChanged -= ConfigServiceLanguageChanged;
            }
        }
    }
}
