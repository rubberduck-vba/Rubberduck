using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using Path = System.IO.Path;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.CodeAnalysis.Inspections.Attributes;
using Rubberduck.CodeAnalysis.Settings;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Resources;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Logistics
{
    internal class Inspector : IInspector
    {
        private const int _maxDegreeOfInspectionParallelism = -1;
        private readonly IConfigurationService<CodeInspectionSettings> _configService;
        private readonly List<IInspection> _inspections;

        public Inspector(IConfigurationService<CodeInspectionSettings> configService, IInspectionProvider inspectionProvider)
        {
            _inspections = inspectionProvider.Inspections.ToList();

            _configService = configService;
            configService.SettingsChanged += ConfigServiceSettingsChanged;
        }

        private void ConfigServiceSettingsChanged(object sender, EventArgs e)
        {
            var config = _configService.Read();
            UpdateInspectionSeverity(config);
        }

        private void UpdateInspectionSeverity(CodeInspectionSettings config)
        {
            foreach (var inspection in _inspections)
            {
                foreach (var setting in config.CodeInspections)
                {
                    if (inspection.Name == setting.Name)
                    {
                        inspection.Severity = setting.Severity;
                        break;
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
            token.ThrowIfCancellationRequested();

            state.OnStatusMessageUpdate(CodeAnalysisUI.CodeInspections_Inspecting);
            var allIssues = new ConcurrentBag<IInspectionResult>();
            token.ThrowIfCancellationRequested();

            var config = _configService.Read();
            UpdateInspectionSeverity(config);
            token.ThrowIfCancellationRequested();

            var parseTreeInspections = _inspections
                .Where(inspection => inspection.Severity != CodeInspectionSeverity.DoNotShow)
                .OfType<IParseTreeInspection>()
                .ToArray();
            token.ThrowIfCancellationRequested();

            foreach (var listener in parseTreeInspections.Select(inspection => inspection.Listener))
            {
                listener.ClearContexts();
            }

            // Prepare ParseTreeWalker based inspections
            var passes = Enum.GetValues(typeof(CodeKind)).Cast<CodeKind>();
            foreach (var parsePass in passes)
            {
                try
                {
                    WalkTrees(config, state, parseTreeInspections.Where(i => i.TargetKindOfCode == parsePass), parsePass);
                }
                catch (Exception e)
                {
                    LogManager.GetCurrentClassLogger().Warn(e);
                }
            }
            token.ThrowIfCancellationRequested();

            var inspectionsToRun = _inspections.Where(inspection =>
                inspection.Severity != CodeInspectionSeverity.DoNotShow &&
                RequiredLibrariesArePresent(inspection, state) &&
                RequiredHostIsPresent(inspection));

            token.ThrowIfCancellationRequested();

            try
            {
                await Task.Run(() => RunInspectionsInParallel(inspectionsToRun, allIssues, token));
            }
            catch (AggregateException exception)
            {
                if (exception.Flatten().InnerExceptions.All(ex => ex is OperationCanceledException))
                {
                    //This eliminates the stack trace, but for the cancellation, this is irrelevant.
                    throw exception.InnerException ?? exception;
                }

                LogManager.GetCurrentClassLogger().Error(exception);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception e)
            {
                LogManager.GetCurrentClassLogger().Error(e);
            }

            // should be "Ready"
            state.OnStatusMessageUpdate(RubberduckUI.ResourceManager.GetString("ParserState_" + state.Status, CultureInfo.CurrentUICulture));
            return allIssues;
        }

        private static bool RequiredLibrariesArePresent(IInspection inspection, RubberduckParserState state)
        {
            var requiredLibraries = inspection.GetType().GetCustomAttributes<RequiredLibraryAttribute>().ToArray();

            if (!requiredLibraries.Any())
            {
                return true;
            }

            var projectNames = state.DeclarationFinder.Projects.Where(project => !project.IsUserDefined).Select(project => project.ProjectName).ToArray();

            return requiredLibraries.All(library => projectNames.Contains(library.LibraryName));
        }

        private static bool RequiredHostIsPresent(IInspection inspection)
        {
            var requiredHost = inspection.GetType().GetCustomAttribute<RequiredHostAttribute>();

            return requiredHost == null || requiredHost.HostNames.Contains(Path.GetFileName(Application.ExecutablePath).ToUpper());
        }

        private static void RunInspectionsInParallel(IEnumerable<IInspection> inspectionsToRun,
            ConcurrentBag<IInspectionResult> allIssues, CancellationToken token)
        {
            var options = new ParallelOptions
            {
                CancellationToken = token,
                MaxDegreeOfParallelism = _maxDegreeOfInspectionParallelism
            };

            Parallel.ForEach(inspectionsToRun,
                options,
                inspection => RunInspection(inspection, allIssues, token)
            );
        }

        private static void RunInspection(IInspection inspection, ConcurrentBag<IInspectionResult> allIssues, CancellationToken token)
        {
            try
            {
                var inspectionResults = inspection.GetInspectionResults(token);

                token.ThrowIfCancellationRequested();

                foreach (var inspectionResult in inspectionResults)
                {
                    allIssues.Add(inspectionResult);
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception e)
            {
                LogManager.GetCurrentClassLogger().Warn(e);
            }
        }

        private void WalkTrees(CodeInspectionSettings settings, RubberduckParserState state, IEnumerable<IParseTreeInspection> inspections, CodeKind codeKind)
        {
            var listeners = inspections
                .Where(i => i.Severity != CodeInspectionSeverity.DoNotShow
                    && i.TargetKindOfCode == codeKind
                    && !IsDisabled(settings, i))
                .Select(inspection => inspection.Listener)
                .ToList();

            if (!listeners.Any())
            {
                return;
            }

            List<KeyValuePair<QualifiedModuleName, IParseTree>> trees;
            switch (codeKind)
            {
                case CodeKind.AttributesCode:
                    trees = state.AttributeParseTrees;
                    break;
                case CodeKind.CodePaneCode:
                    trees = state.ParseTrees;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(codeKind), codeKind, null);
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
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private bool _isDisposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_isDisposed || !disposing)
            {
                return;
            }

            if (_configService != null)
            {
                _configService.SettingsChanged -= ConfigServiceSettingsChanged;
            }

            _inspections.Clear();
            _isDisposed = true;
        }
    }
}