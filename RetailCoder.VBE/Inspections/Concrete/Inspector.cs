using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI;
using Rubberduck.Inspections.Results;

namespace Rubberduck.Inspections.Concrete
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

                // Prepare ParseTreeWalker based inspections
                var parseTreeWalkResults = GetParseTreeResults(config, state);
                foreach (var parseTreeInspection in _inspections.OfType<IParseTreeInspection>().Where(inspection => inspection.Severity != CodeInspectionSeverity.DoNotShow))
                {
                    parseTreeInspection.SetResults(parseTreeWalkResults);
                }

                var inspections = _inspections.Where(inspection => inspection.Severity != CodeInspectionSeverity.DoNotShow)
                    .Select(inspection =>
                        Task.Run(() =>
                        {
                            token.ThrowIfCancellationRequested();
                            var inspectionResults = inspection.GetInspectionResults();
                            
                            foreach (var inspectionResult in inspectionResults)
                            {
                                allIssues.Add(inspectionResult);
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

                var issuesByType = allIssues.GroupBy(issue => issue.GetType())
                                            .ToDictionary(grouping => grouping.Key, grouping => grouping.ToList());
                var results = issuesByType.Where(kv => kv.Value.Count <= AGGREGATION_THRESHOLD)
                    .SelectMany(kv => kv.Value)
                    .Union(issuesByType.Where(kv => kv.Value.Count > AGGREGATION_THRESHOLD)
                    .Select(kv => new AggregateInspectionResult(kv.Value.OrderBy(i => i.QualifiedSelection).First(), kv.Value.Count)))
                    .ToList();

                state.OnStatusMessageUpdate(RubberduckUI.ResourceManager.GetString("ParserState_" + state.Status, CultureInfo.CurrentUICulture)); // should be "Ready"
                return results;
            }

            private IReadOnlyList<QualifiedContext> GetParseTreeResults(Configuration config, RubberduckParserState state)
            {
                var result = new List<QualifiedContext>();
                var settings = config.UserSettings.CodeInspectionSettings;

                foreach (var componentTreePair in state.ParseTrees)
                {
                    /*
                    Need to reinitialize these for each and every ParseTree we process, since the results are aggregated in the instances themselves 
                    before moving them into the ParseTreeResults after qualifying them 
                    */
                    var obsoleteCallStatementListener = IsDisabled<ObsoleteCallStatementInspection>(settings) ? null : new ObsoleteCallStatementInspection.ObsoleteCallStatementListener();
                    var obsoleteLetStatementListener = IsDisabled<ObsoleteLetStatementInspection>(settings) ? null : new ObsoleteLetStatementInspection.ObsoleteLetStatementListener();
                    var obsoleteCommentSyntaxListener = IsDisabled<ObsoleteCommentSyntaxInspection>(settings) ? null : new ObsoleteCommentSyntaxInspection.ObsoleteCommentSyntaxListener();
                    var emptyStringLiteralListener = IsDisabled<EmptyStringLiteralInspection>(settings) ? null : new EmptyStringLiteralInspection.EmptyStringLiteralListener();
                    var argListWithOneByRefParamListener = IsDisabled<ProcedureCanBeWrittenAsFunctionInspection>(settings) ? null : new ProcedureCanBeWrittenAsFunctionInspection.SingleByRefParamArgListListener();
                    var invalidAnnotationListener = IsDisabled<MissingAnnotationArgumentInspection>(settings) ? null : new MissingAnnotationArgumentInspection.InvalidAnnotationStatementListener();

                    var combinedListener = new CombinedParseTreeListener(new IParseTreeListener[]{
                        obsoleteCallStatementListener,
                        obsoleteLetStatementListener,
                        obsoleteCommentSyntaxListener,
                        emptyStringLiteralListener,
                        argListWithOneByRefParamListener,
                        invalidAnnotationListener
                    });

                    ParseTreeWalker.Default.Walk(combinedListener, componentTreePair.Value);

                    if (argListWithOneByRefParamListener != null)
                    {
                        result.AddRange(argListWithOneByRefParamListener.Contexts.Select(context => new QualifiedContext<VBAParser.ArgListContext>(componentTreePair.Key, context)));
                    }
                    if (emptyStringLiteralListener != null)
                    {
                        result.AddRange(emptyStringLiteralListener.Contexts.Select(context => new QualifiedContext<VBAParser.LiteralExpressionContext>(componentTreePair.Key, context)));
                    }
                    if (obsoleteLetStatementListener != null)
                    {
                        result.AddRange(obsoleteLetStatementListener.Contexts.Select(context => new QualifiedContext<VBAParser.LetStmtContext>(componentTreePair.Key, context)));
                    }
                    if (obsoleteCommentSyntaxListener != null)
                    {
                        result.AddRange(obsoleteCommentSyntaxListener.Contexts.Select(context => new QualifiedContext<VBAParser.RemCommentContext>(componentTreePair.Key, context)));
                    }
                    if (obsoleteCallStatementListener != null)
                    {
                        result.AddRange(obsoleteCallStatementListener.Contexts.Select(context => new QualifiedContext<VBAParser.CallStmtContext>(componentTreePair.Key, context)));
                    }
                    if (invalidAnnotationListener != null)
                    {
                        result.AddRange(invalidAnnotationListener.Contexts.Select(context => new QualifiedContext<VBAParser.AnnotationContext>(componentTreePair.Key, context)));
                    }
                }
                return result;
            }

            private bool IsDisabled<TInspection>(CodeInspectionSettings config) where TInspection : IInspection
            {
                var setting = config.GetSetting<TInspection>();
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
