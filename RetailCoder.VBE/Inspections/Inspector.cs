using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;

namespace Rubberduck.Inspections
{
    public interface IInspectorFactory
    {
        IInspector Create();
    }

    public class Inspector : IInspector, IDisposable
    {
        private readonly IGeneralConfigService _configService;
        private readonly IEnumerable<IInspection> _inspections;

        public Inspector(IGeneralConfigService configService, IEnumerable<IInspection> inspections)
        {
            _inspections = inspections;

            _configService = configService;
            configService.SettingsChanged += ConfigServiceSettingsChanged;
            UpdateInspectionSeverity();
        }

        private void ConfigServiceSettingsChanged(object sender, EventArgs e)
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
                        var inspectionResults = inspection.GetInspectionResults();
                        var results = inspectionResults as IList<CodeInspectionResultBase> ?? inspectionResults.ToList();

                        if (results.Any())
                        {
                            //OnIssuesFound(results);

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

        public void Dispose()
        {
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposing) { return; }

            if (_configService != null)
            {
                _configService.SettingsChanged -= ConfigServiceSettingsChanged;
            }
        }
    }
}
