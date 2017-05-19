using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI;
using Rubberduck.UI.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    namespace Rubberduck.Inspections
    {
        public class Inspector : IInspector
        {
            private readonly IGeneralConfigService _configService;
            private readonly List<IInspection> _inspections;
            private readonly int AGGREGATION_THRESHOLD = 128;

            public Inspector(IGeneralConfigService configService, IEnumerable<IInspection> inspections)
            {
                _inspections = inspections.ToList();

                _configService = configService;
                configService.SettingsChanged += ConfigServiceSettingsChanged;
            }

            private void ConfigServiceSettingsChanged(object sender, EventArgs e)
            {
                var config = _configService.LoadConfiguration();
                UpdateInspectionSeverity(config);
            }

            private void UpdateInspectionSeverity(Configuration config)
            {
                foreach (var inspection in _inspections)
                {
                    foreach (var setting in config.UserSettings.CodeInspectionSettings.CodeInspections)
                    {
                        if (inspection.Name == setting.Name)
                        {
                            inspection.Severity = setting.Severity;
                        }
                    }
                }
            }

            public async Task<IEnumerable<IInspectionResult>> FindIssuesAsync(RubberduckParserState state, CancellationToken token)
            {
                if (state == null || !state.AllUserDeclarations.Any())
                {
                    return new IInspectionResult[] { };
                }

                state.OnStatusMessageUpdate(RubberduckUI.CodeInspections_Inspecting);

                var config = _configService.LoadConfiguration();
                UpdateInspectionSeverity(config);

                var allIssues = new ConcurrentBag<IInspectionResult>();

                var parseTreeInspections = _inspections.OfType<IParseTreeInspection>().ToArray();

                // Prepare ParseTreeWalker based inspections
                var passes = Enum.GetValues(typeof (ParsePass)).Cast<ParsePass>();
                foreach (var parsePass in passes)
                {
                    try
                    {
                        WalkTrees(config.UserSettings.CodeInspectionSettings, state, parseTreeInspections.Where(i => i.Pass == parsePass), parsePass);
                    }
                    catch (Exception e)
                    {
                        LogManager.GetCurrentClassLogger().Warn(e);
                    }
                }

                var inspections = _inspections.Where(inspection => inspection.Severity != CodeInspectionSeverity.DoNotShow)
                    .Select(inspection =>
                        Task.Run(() =>
                        {
                            token.ThrowIfCancellationRequested();
                            try
                            {
                                var inspectionResults = inspection.GetInspectionResults();
                            
                                foreach (var inspectionResult in inspectionResults)
                                {
                                    allIssues.Add(inspectionResult);
                                }
                            }
                            catch(Exception e)
                            {
                                LogManager.GetCurrentClassLogger().Warn(e);
                            }
                        }, token)).ToList();

                try
                {
                    await Task.WhenAll(inspections);
                }
                catch (Exception e)
                {
                    LogManager.GetCurrentClassLogger().Error(e);
                }

                var issuesByType = allIssues.GroupBy(issue => issue.Inspection.Name)
                                            .ToDictionary(grouping => grouping.Key, grouping => grouping.ToList());
                var results = issuesByType.Where(kv => kv.Value.Count <= AGGREGATION_THRESHOLD)
                    .SelectMany(kv => kv.Value)
                    .Union(issuesByType.Where(kv => kv.Value.Count > AGGREGATION_THRESHOLD)
                    .Select(kv => new AggregateInspectionResult(kv.Value.OrderBy(i => i.QualifiedSelection).First(), kv.Value.Count)))
                    .ToList();

                state.OnStatusMessageUpdate(RubberduckUI.ResourceManager.GetString("ParserState_" + state.Status, CultureInfo.CurrentUICulture)); // should be "Ready"
                return results;
            }

            private void WalkTrees(CodeInspectionSettings settings, RubberduckParserState state, IEnumerable<IParseTreeInspection> inspections, ParsePass pass)
            {
                var listeners =
                    inspections.Where(i => i.Pass == pass && i.Severity != CodeInspectionSeverity.DoNotShow && !IsDisabled(settings, i))
                        .Select(inspection => inspection.Listener)
                        .ToList();

                List<KeyValuePair<QualifiedModuleName, IParseTree>> trees;
                switch (pass)
                {
                    case ParsePass.AttributesPass:
                        trees = state.AttributeParseTrees;
                        break;
                    case ParsePass.CodePanePass:
                        trees = state.ParseTrees;
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(pass), pass, null);
                }

                foreach (var componentTreePair in trees)
                {
                    foreach (var listener in listeners)
                    {
                        listener.CurrentModuleName = componentTreePair.Key;
                    }

                    ParseTreeWalker.Default.Walk(new CombinedParseTreeListener(listeners), componentTreePair.Value);
                }
            }

            private bool IsDisabled(CodeInspectionSettings config, IInspection inspection)
            {
                var setting = config.GetSetting(inspection.GetType());
                return setting != null && setting.Severity == CodeInspectionSeverity.DoNotShow;
            }

            public void Dispose()
            {
                if (_configService != null)
                {
                    _configService.SettingsChanged -= ConfigServiceSettingsChanged;
                }

                _inspections.Clear();
            }
        }
    }
}
