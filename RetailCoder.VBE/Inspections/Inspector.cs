using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    namespace Rubberduck.Inspections
    {
        public class Inspector : IInspector
        {
            private readonly IGeneralConfigService _configService;
            private readonly List<IInspection> _inspections;

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
                        if (inspection.Description == setting.Description)
                        {
                            inspection.Severity = setting.Severity;
                        }
                    }
                }
            }

            public async Task<IEnumerable<ICodeInspectionResult>> FindIssuesAsync(RubberduckParserState state, CancellationToken token)
            {
                if (state == null || !state.AllUserDeclarations.Any())
                {
                    return new ICodeInspectionResult[] { };
                }

                state.OnStatusMessageUpdate(RubberduckUI.CodeInspections_Inspecting);

                var config = _configService.LoadConfiguration();
                UpdateInspectionSeverity(config);

                var allIssues = new ConcurrentBag<ICodeInspectionResult>();

                // Prepare ParseTreeWalker based inspections
                var parseTreeWalkResults = GetParseTreeResults(config, state);
                foreach (var parseTreeInspection in _inspections.Where(inspection => inspection.Severity != CodeInspectionSeverity.DoNotShow && inspection is IParseTreeInspection))
                {
                    (parseTreeInspection as IParseTreeInspection).ParseTreeResults = parseTreeWalkResults;
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

                await Task.WhenAll(inspections);
                state.OnStatusMessageUpdate(RubberduckUI.ResourceManager.GetString("ParserState_" + state.Status, UI.Settings.Settings.Culture)); // should be "Ready"
                return allIssues;
            }

            private ParseTreeResults GetParseTreeResults(Configuration config, RubberduckParserState state)
            {
                var result = new ParseTreeResults();
                var settings = config.UserSettings.CodeInspectionSettings;

                foreach (var componentTreePair in state.ParseTrees)
                {
                    /*
                    Need to reinitialize these for each and every ParseTree we process, since the results are aggregated in the instances themselves 
                    before moving them into the ParseTreeResults after qualifying them 
                    */
                    var obsoleteCallStatementListener = IsDisabled<ObsoleteCallStatementInspection>(settings) ? null : new ObsoleteCallStatementInspection.ObsoleteCallStatementListener();
                    var obsoleteLetStatementListener = IsDisabled<ObsoleteLetStatementInspection>(settings) ? null : new ObsoleteLetStatementInspection.ObsoleteLetStatementListener();
                    var emptyStringLiteralListener = IsDisabled<EmptyStringLiteralInspection>(settings) ? null : new EmptyStringLiteralInspection.EmptyStringLiteralListener();
                    var argListWithOneByRefParamListener = IsDisabled<ProcedureCanBeWrittenAsFunctionInspection>(settings) ? null : new ProcedureCanBeWrittenAsFunctionInspection.ArgListWithOneByRefParamListener();
                    var malformedAnnotationListenter = IsDisabled<MalformedAnnotationInspection>(settings) ? null : new MalformedAnnotationInspection.MalformedAnnotationStatementListener();

                    var combinedListener = new CombinedParseTreeListener(new IParseTreeListener[]{
                        obsoleteCallStatementListener,
                        obsoleteLetStatementListener,
                        emptyStringLiteralListener,
                        argListWithOneByRefParamListener,
                        malformedAnnotationListenter
                    });

                    ParseTreeWalker.Default.Walk(combinedListener, componentTreePair.Value);

                    result.ArgListsWithOneByRefParam = argListWithOneByRefParamListener == null 
                        ? Enumerable.Empty<QualifiedContext>() 
                        : result.ArgListsWithOneByRefParam.Concat(argListWithOneByRefParamListener.Contexts.Select(context => new QualifiedContext(componentTreePair.Key, context)));
                    result.EmptyStringLiterals = emptyStringLiteralListener == null 
                        ? Enumerable.Empty<QualifiedContext>() 
                        : result.EmptyStringLiterals.Concat(emptyStringLiteralListener.Contexts.Select(context => new QualifiedContext(componentTreePair.Key, context)));
                    result.ObsoleteLetContexts = obsoleteLetStatementListener == null 
                        ? Enumerable.Empty<QualifiedContext>() 
                        : result.ObsoleteLetContexts.Concat(obsoleteLetStatementListener.Contexts.Select(context => new QualifiedContext(componentTreePair.Key, context)));
                    result.ObsoleteCallContexts = obsoleteCallStatementListener == null 
                        ? Enumerable.Empty<QualifiedContext>() 
                        : result.ObsoleteCallContexts.Concat(obsoleteCallStatementListener.Contexts.Select(context => new QualifiedContext(componentTreePair.Key, context)));
                    result.MalformedAnnotations = malformedAnnotationListenter == null 
                        ? Enumerable.Empty<QualifiedContext>() 
                        : result.MalformedAnnotations.Concat(malformedAnnotationListenter.Contexts.Select(context => new QualifiedContext<VBAParser.AnnotationContext>(componentTreePair.Key, context)));
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
