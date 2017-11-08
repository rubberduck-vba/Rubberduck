using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System.Diagnostics;
using System.Linq;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA
{
    public class ParseCoordinator : IParseCoordinator
    {
        public RubberduckParserState State { get; }

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IParsingStageService _parsingStageService;
        private readonly IProjectManager _projectManager;
        private readonly IParsingCacheService _parsingCacheService;
        private readonly IParserStateManager _parserStateManager;

        private readonly bool _isTestScope;

        public ParseCoordinator(
            RubberduckParserState state,
            IParsingStageService parsingStageService,
            IParsingCacheService parsingCacheService,
            IProjectManager projectManager,
            IParserStateManager parserStateManager,
            bool isTestScope = false)
        {
            if (state == null)
            {
                throw new ArgumentNullException(nameof(state));
            }
            if (parsingStageService == null)
            {
                throw new ArgumentNullException(nameof(parsingStageService));
            }
            if (parsingStageService == null)
            {
                throw new ArgumentNullException(nameof(parsingStageService));
            }
            if (parsingCacheService == null)
            {
                throw new ArgumentNullException(nameof(parsingCacheService));
            }
            if (parserStateManager == null)
            {
                throw new ArgumentNullException(nameof(parserStateManager));
            }

            State = state;
            _parsingStageService = parsingStageService;
            _projectManager = projectManager;
            _parsingCacheService = parsingCacheService;
            _parserStateManager = parserStateManager;
            _isTestScope = isTestScope;

            state.ParseRequest += ReparseRequested;
        }

        // Do not access this from anywhere but ReparseRequested.
        // ReparseRequested needs to have a reference to the cancellation token.
        private CancellationTokenSource _currentCancellationTokenSource = new CancellationTokenSource();

        private readonly object _cancellationSyncObject = new object();
        private readonly object _parsingRunSyncObject = new object();

        private void ReparseRequested(object sender, EventArgs e)
        {
            CancellationToken token;
            lock (_cancellationSyncObject)
            {
                Cancel();

                if (_currentCancellationTokenSource == null)
                {
                    Logger.Error("Tried to request a parse after the final cancellation.");
                    return;
                }

                token = _currentCancellationTokenSource.Token;
            }


            if (!_isTestScope)
            {
                Task.Run(() => ParseAll(sender, token), token);
            }
            else
            {
                ParseInternal(token);
            }
        }

        private void Cancel(bool createNewTokenSource = true)
        {
            lock (_cancellationSyncObject)
            {
                if (_currentCancellationTokenSource == null)
                {
                    Logger.Error("Tried to cancel a parse after the final cancellation.");
                    return;
                }

                var oldTokenSource = _currentCancellationTokenSource;
                _currentCancellationTokenSource = createNewTokenSource ? new CancellationTokenSource() : null;

                oldTokenSource.Cancel();
                oldTokenSource.Dispose();
            }
        }

        /// <summary>
        /// For the use of tests only
        /// </summary>
        /// 
        public void Parse(CancellationTokenSource token)
        {
            SetSavedCancellationTokenSource(token);
            ParseInternal(token.Token);
        }

        /// <summary>
        /// For the use of tests only
        /// </summary>
        /// 
        private void SetSavedCancellationTokenSource(CancellationTokenSource tokenSource)
        {
            var oldTokenSource = _currentCancellationTokenSource;
            _currentCancellationTokenSource = tokenSource;

            oldTokenSource?.Cancel();
            oldTokenSource?.Dispose();
        }

        private void ParseInternal(CancellationToken token)
        {
            var lockTaken = false;
            try
            {
                Monitor.Enter(_parsingRunSyncObject, ref lockTaken);
                ParseAllInternal(this, token);
            }
            catch (OperationCanceledException)
            {
                //This is the point to which the cancellation should break.
            }
            finally
            {
                if (lockTaken) Monitor.Exit(_parsingRunSyncObject);
            }
        }


        private void ExecuteCommonParseActivities(IReadOnlyCollection<QualifiedModuleName> toParse, IReadOnlyCollection<QualifiedModuleName> toReresolveReferencesInput, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            var toReresolveReferences = new HashSet<QualifiedModuleName>();
            toReresolveReferences.UnionWith(toReresolveReferencesInput);
            token.ThrowIfCancellationRequested();

            _parserStateManager.SetModuleStates(toParse, ParserState.Pending, token);
            token.ThrowIfCancellationRequested();

            _parserStateManager.SetStatusAndFireStateChanged(this, ParserState.LoadingReference, token);
            token.ThrowIfCancellationRequested();

            _parsingStageService.SyncComReferences(State.Projects, token);
            if (_parsingStageService.LastSyncOfCOMReferencesLoadedReferences || _parsingStageService.COMReferencesUnloadedUnloadedInLastSync.Any())
            {
                var unloadedReferences = _parsingStageService.COMReferencesUnloadedUnloadedInLastSync;
                var additionalModulesToBeReresolved = OtherModulesReferencingAnyNotToBeParsed(unloadedReferences.ToHashSet().AsReadOnly(), toParse);
                toReresolveReferences.UnionWith(additionalModulesToBeReresolved);
                _parserStateManager.SetModuleStates(additionalModulesToBeReresolved, ParserState.ResolvingReferences, token);
                ClearModuleToModuleReferences(unloadedReferences);
                RefreshDeclarationFinder();
            }
            token.ThrowIfCancellationRequested();

            _parsingStageService.LoadBuitInDeclarations();
            if (_parsingStageService.LastLoadOfBuiltInDeclarationsLoadedDeclarations)
            { 
                RefreshDeclarationFinder();
            }
            token.ThrowIfCancellationRequested();

            IReadOnlyCollection<QualifiedModuleName> toResolveReferences;
            if (!toParse.Any())
            {
                toResolveReferences = toReresolveReferences.AsReadOnly();
            }
            else
            {
                toResolveReferences = ModulesForWhichToResolveReferences(toParse, toReresolveReferences);
                token.ThrowIfCancellationRequested();

                PerformPreParseCleanup(toResolveReferences, token);
                token.ThrowIfCancellationRequested();

                _parserStateManager.SetModuleStates(toParse, ParserState.Parsing, token);
                token.ThrowIfCancellationRequested();

                _parsingStageService.ParseModules(toParse, token);

                if (token.IsCancellationRequested || State.Status >= ParserState.Error)
                {
                    throw new OperationCanceledException(token);
                }

                _parserStateManager.EvaluateOverallParserState(token);

                if (token.IsCancellationRequested || State.Status >= ParserState.Error)
                {
                    throw new OperationCanceledException(token);
                }

                _parserStateManager.SetModuleStates(toParse, ParserState.ResolvingDeclarations, token);
                token.ThrowIfCancellationRequested();

                _parsingStageService.ResolveDeclarations(toParse, token);
            }

            if (token.IsCancellationRequested || State.Status >= ParserState.Error)
            {
                throw new OperationCanceledException(token);
            }

            //We need to refresh the DeclarationFinder before the handlers for ResolvedDeclarations run no matter 
            //whether we parsed or resolved something because modules not referenced by any remeining module might
            //have been removed. E.g. the CodeExplorer needs this update. 
            RefreshDeclarationFinder();
            token.ThrowIfCancellationRequested();

            //Explicitly setting the overall state here guarantees that the handlers attached 
            //to the state change to ResolvedDeclarations always run, provided there is no error. 
            State.SetStatusAndFireStateChanged(this, ParserState.ResolvedDeclarations);

            if (token.IsCancellationRequested || State.Status >= ParserState.Error)
            {
                throw new OperationCanceledException(token);
            }

            _parserStateManager.SetModuleStates(toResolveReferences, ParserState.ResolvingReferences, token);
            token.ThrowIfCancellationRequested();

            _parsingStageService.ResolveReferences(toResolveReferences, token);

            if (token.IsCancellationRequested || State.Status >= ParserState.Error)
            {
                throw new OperationCanceledException(token);
            }

            RefreshDeclarationFinder();
            token.ThrowIfCancellationRequested();

            //At this point all modules should either be in the Ready state or in an error state.
            //This is the point where the change of the overall state to Ready is triggered on the success path.
            _parserStateManager.EvaluateOverallParserState(token);
            token.ThrowIfCancellationRequested();
        }

        private void ClearModuleToModuleReferences(IEnumerable<QualifiedModuleName> modules)
        {
            foreach (var module in modules)
            {
                _parsingCacheService.ClearModuleToModuleReferencesToModule(module);
                _parsingCacheService.ClearModuleToModuleReferencesFromModule(module);
            }
        }

        private void PerformPreParseCleanup(IReadOnlyCollection<QualifiedModuleName> toResolveReferences, CancellationToken token)
        {
            _parsingCacheService.ClearSupertypes(toResolveReferences);
            //This is purely a security measure. In the success path, the reference remover removes the referernces.
            _parsingCacheService.RemoveReferencesBy(toResolveReferences, token);

        }

        private void RefreshDeclarationFinder()
        {
            _parsingCacheService.RefreshDeclarationFinder();
        }

        private IReadOnlyCollection<QualifiedModuleName> ModulesForWhichToResolveReferences(IReadOnlyCollection<QualifiedModuleName> modulesToParse, IEnumerable<QualifiedModuleName> toReresolveReferences)
        {
            var toResolveReferences = modulesToParse.ToHashSet();
            toResolveReferences.UnionWith(_parsingCacheService.ModulesReferencingAny(modulesToParse));
            toResolveReferences.UnionWith(toReresolveReferences);
            return toResolveReferences.AsReadOnly();
        }


        /// <summary>
        /// Starts parsing all components of all unprotected VBProjects associated with the VBE-Instance passed to the constructor of this parser instance.
        /// </summary>
        private void ParseAll(object requestor, CancellationToken token)
        {
            Stopwatch watch = null;
            var lockTaken = false;
            try
            {
                Monitor.Enter(_parsingRunSyncObject, ref lockTaken);
                
                watch = Stopwatch.StartNew();
                Logger.Debug("Parsing run started. (thread {0}).", Thread.CurrentThread.ManagedThreadId);
                
                ParseAllInternal(requestor, token);
            }
            catch (OperationCanceledException)
            {
                //This is the point to which the cancellation should break.
                Logger.Debug("Parsing run got canceled. (thread {0}).", Thread.CurrentThread.ManagedThreadId);
            }
            catch (Exception exception)
            {
                Logger.Error(exception, "Unexpected exception thrown in parsing run. (thread {0}).", Thread.CurrentThread.ManagedThreadId);
                if (!(_parserStateManager.OverallParserState >= ParserState.Error))
                {
                    _parserStateManager.SetStatusAndFireStateChanged(this, ParserState.Error, token);
                }
            }
            finally
            {
                if (watch != null && watch.IsRunning) watch.Stop();
                if (lockTaken) Monitor.Exit(_parsingRunSyncObject);
            }
            if (watch != null) Logger.Debug("Parsing run finished after {0}s. (thread {1}).", watch.Elapsed.TotalSeconds, Thread.CurrentThread.ManagedThreadId);
        }

        private void ParseAllInternal(object requestor, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            _parserStateManager.SetStatusAndFireStateChanged(requestor, ParserState.Pending, token);
            token.ThrowIfCancellationRequested();

            _projectManager.RefreshProjects();
            token.ThrowIfCancellationRequested();

            var modules = _projectManager.AllModules();
            token.ThrowIfCancellationRequested();

            var toParse = modules.Where(module => State.IsNewOrModified(module)).ToHashSet();
            token.ThrowIfCancellationRequested();

            toParse.UnionWith(modules.Where(module => _parserStateManager.GetModuleState(module) != ParserState.Ready));
            token.ThrowIfCancellationRequested();

            var removedModules = RemovedModules(modules);
            token.ThrowIfCancellationRequested();

            var removedProjects = RemovedProjects(_projectManager.Projects);
            token.ThrowIfCancellationRequested();

            removedModules.UnionWith(ModulesInProjects(removedProjects));
            token.ThrowIfCancellationRequested();

            var toReResolveReferences = OtherModulesReferencingAnyNotToBeParsed(removedModules.AsReadOnly(), toParse.AsReadOnly());
            token.ThrowIfCancellationRequested();

            _parserStateManager.SetModuleStates(toReResolveReferences, ParserState.ResolvingReferences, token);
            token.ThrowIfCancellationRequested();

            CleanUpRemovedComponents(removedModules.AsReadOnly(), token);
            token.ThrowIfCancellationRequested();

            //This must come after the component cleanup because of cache invalidation.
            CleanUpRemovedProjects(removedProjects);
            token.ThrowIfCancellationRequested();

            ExecuteCommonParseActivities(toParse.AsReadOnly(), toReResolveReferences, token);
        }

        private IReadOnlyCollection<QualifiedModuleName> OtherModulesReferencingAnyNotToBeParsed(IReadOnlyCollection<QualifiedModuleName> removedModules, IReadOnlyCollection<QualifiedModuleName> toParse)
        {
            return _parsingCacheService.ModulesReferencingAny(removedModules)
                        .Where(qmn => !removedModules.Contains(qmn) && !toParse.Contains(qmn))
                        .ToHashSet().AsReadOnly();
        }

        private IEnumerable<QualifiedModuleName> ModulesInProjects(IReadOnlyCollection<string> removedProjects)
        {
            return State.DeclarationFinder.UserDeclarations(DeclarationType.Module)
                    .Where(declaration => removedProjects.Contains(declaration.ProjectId))
                    .Select(declaration => declaration.QualifiedName.QualifiedModuleName);
        }

        private void CleanUpRemovedComponents(IReadOnlyCollection<QualifiedModuleName> removedModules, CancellationToken token)
        {
            if (removedModules.Any())
            {
                _parsingCacheService.RemoveReferencesBy(removedModules, token);
                _parsingCacheService.ClearSupertypes(removedModules);
                ClearModuleToModuleReferences(removedModules);
                ClearStateCache(removedModules);
            }
        }

        private void ClearStateCache(IEnumerable<QualifiedModuleName> modules)
        {
            foreach (var module in modules)
            {
                State.ClearStateCache(module);
            }
        }

        private void CleanUpRemovedProjects(IReadOnlyCollection<string> removedProjectIds)
        {
            ClearStateCache(removedProjectIds);
        }

        private void ClearStateCache(IEnumerable<string> projectIds)
        {
            foreach (var projectId in projectIds)
            {
                State.ClearStateCache(projectId);
            }
        }

        private HashSet<QualifiedModuleName> RemovedModules(IReadOnlyCollection<QualifiedModuleName> modules)
        {
            var modulesWithModuleDeclarations = State.DeclarationFinder.UserDeclarations(DeclarationType.Module).Select(declaration => declaration.QualifiedName.QualifiedModuleName);
            var currentlyExistingModules = modules.ToHashSet();
            var removedModuledecalrations = modulesWithModuleDeclarations.Where(module => !currentlyExistingModules.Contains(module));
            return removedModuledecalrations.ToHashSet();
        }

        private IReadOnlyCollection<string> RemovedProjects(IReadOnlyCollection<IVBProject> projects)
        {
            var projectsWithProjectDeclarations = State.DeclarationFinder.UserDeclarations(DeclarationType.Project).Select(declaration => new Tuple<string,string>(declaration.ProjectId, declaration.ProjectName));
            var currentlyExistingProjects = projects.Select(project => new Tuple<string, string>(project.ProjectId, project.Name)).ToHashSet();
            var removedProjects = projectsWithProjectDeclarations.Where(project => !currentlyExistingProjects.Contains(project));
            return removedProjects.Select(tuple => tuple.Item1).ToHashSet().AsReadOnly();
        }


        public void Dispose()
        {
            State.ParseRequest -= ReparseRequested;
            Cancel(false);
        }
    }
}