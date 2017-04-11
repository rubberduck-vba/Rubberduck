using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System.Diagnostics;
using System.Linq;
using NLog;


namespace Rubberduck.Parsing.VBA
{
    public class ParseCoordinator : IParseCoordinator
    {
        public RubberduckParserState State { get { return _state; } }

        private readonly RubberduckParserState _state;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IParsingStageService _parsingStageService;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IProjectManager _projectManager;
        private readonly IModuleToModuleReferenceManager _moduleToModuleReferenceManager;
        private readonly IParserStateManager _parserStateManager;
        private readonly IReferenceRemover _referenceRemover;

        private readonly bool _isTestScope;

        public ParseCoordinator(
            RubberduckParserState state,
            IParsingStageService parsingStageService,
            IDeclarationFinderProvider declarationFinderProvider,
            IProjectManager projectManager,
            IModuleToModuleReferenceManager moduleToModuleReferenceManager,
            IParserStateManager parserStateManager,
            IReferenceRemover referenceRemover,
            bool isTestScope = false)
        {
            _state = state;
            _parsingStageService = parsingStageService;
            _declarationFinderProvider = declarationFinderProvider;
            _projectManager = projectManager;
            _moduleToModuleReferenceManager = moduleToModuleReferenceManager;
            _parserStateManager = parserStateManager;
            _referenceRemover = referenceRemover;
            _isTestScope = isTestScope;

            state.ParseRequest += ReparseRequested;
        }

        // Do not access this from anywhere but ReparseRequested.
        // ReparseRequested needs to have a reference to all the cancellation tokens,
        // but the cancelees need to use their own token.
        private readonly List<CancellationTokenSource> _cancellationTokens = new List<CancellationTokenSource> { new CancellationTokenSource() };

        private readonly object _cancellationSyncObject = new object();
        private readonly object _parsingRunSyncObject = new object();

        private void ReparseRequested(object sender, EventArgs e)
        {
            CancellationToken token;
            lock (_cancellationSyncObject)
            {
                Cancel();
                token = _cancellationTokens[0].Token;
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
            lock (_cancellationTokens[0])
            {
                _cancellationTokens[0].Cancel();
                _cancellationTokens[0].Dispose();
                if (createNewTokenSource)
                {
                    _cancellationTokens.Add(new CancellationTokenSource());
                }
                _cancellationTokens.RemoveAt(0);
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

        private void SetSavedCancellationTokenSource(CancellationTokenSource token)
        {
            if (_cancellationTokens.Any())
            {
                _cancellationTokens[0].Cancel();
                _cancellationTokens[0].Dispose();
                _cancellationTokens[0] = token;
            }
            else
            {
                _cancellationTokens.Add(token);
            }
        }

        private void ParseInternal(CancellationToken token)
        {
            var lockTaken = false;
            try
            {
                Monitor.Enter(_parsingRunSyncObject, ref lockTaken);
                ParseInternalInternal(token);
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

        private void ParseInternalInternal(CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            _projectManager.RefreshProjects();
            token.ThrowIfCancellationRequested();

            var modules = _projectManager.AllModules();
            token.ThrowIfCancellationRequested();

            // tests do not fire events when components are removed--clear components
            ClearComponentStateCacheForTests();
            token.ThrowIfCancellationRequested();

            ExecuteCommonParseActivities(modules, token);
        }

        private void ClearComponentStateCacheForTests()
        {
            foreach (var tree in State.ParseTrees)
            {
                State.ClearStateCache(tree.Key);    // handle potentially removed components without crashing
            }
        }

        private void ExecuteCommonParseActivities(IReadOnlyCollection<QualifiedModuleName> toParse, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            
            _parserStateManager.SetModuleStates(toParse, ParserState.Pending, token);
            token.ThrowIfCancellationRequested();

            _parserStateManager.SetStatusAndFireStateChanged(this, ParserState.LoadingReference, token);
            token.ThrowIfCancellationRequested();

            _parsingStageService.SyncComReferences(State.Projects, token);
            if (_parsingStageService.LastSyncOfCOMReferencesLoadedReferences || _parsingStageService.LastSyncOfCOMReferencesUnloadedReferences)
            {
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
                toResolveReferences = new HashSet<QualifiedModuleName>().AsReadOnly();
            }
            else
            {
                toResolveReferences = ModulesForWhichToResolveReferences(toParse);
                PerformPreParseCleanup(toParse, toResolveReferences, token);
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
                RefreshDeclarationFinder();
            }

            if (token.IsCancellationRequested || State.Status >= ParserState.Error)
            {
                throw new OperationCanceledException(token);
            }

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

            _parserStateManager.EvaluateOverallParserState(token);
            token.ThrowIfCancellationRequested();
        }

        private void RefreshDeclarationFinder()
        {
            _declarationFinderProvider.RefreshDeclarationFinder();
        }

        private IReadOnlyCollection<QualifiedModuleName> ModulesForWhichToResolveReferences(IReadOnlyCollection<QualifiedModuleName> modulesToParse)
        {
            var toResolveReferences = modulesToParse.ToHashSet();
            toResolveReferences.UnionWith(_moduleToModuleReferenceManager.ModulesReferencingAny(modulesToParse));
            return toResolveReferences.AsReadOnly();
        }

        private void PerformPreParseCleanup(IReadOnlyCollection<QualifiedModuleName> modulesToParse, IReadOnlyCollection<QualifiedModuleName> toResolveReferences, CancellationToken token)
        {
            _moduleToModuleReferenceManager.ClearModuleToModuleReferencesFromModule(modulesToParse);
            _referenceRemover.RemoveReferencesBy(toResolveReferences, token); 
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
            if (watch != null) Logger.Debug("Parsing run finished after {0}s. (thread {1}).", watch.Elapsed.Seconds, Thread.CurrentThread.ManagedThreadId);
        }


        private void ParseAllInternal(object requestor, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            Thread.Sleep(50); //Simplistic hack to give the VBE time to do its work in case the parsing run is requested from an event handler. 
            token.ThrowIfCancellationRequested();

            _projectManager.RefreshProjects();
            token.ThrowIfCancellationRequested();

            var modules = _projectManager.AllModules();
            token.ThrowIfCancellationRequested();

            var componentsRemoved = CleanUpRemovedComponents(modules, token);
            token.ThrowIfCancellationRequested();

            var toParse = modules.Where(module => State.IsNewOrModified(module)).ToHashSet();
            token.ThrowIfCancellationRequested();

            toParse.UnionWith(modules.Where(module => _parserStateManager.GetModuleState(module) != ParserState.Ready));
            token.ThrowIfCancellationRequested();           

            if (toParse.Count == 0)
            {
                if (componentsRemoved)  // trigger UI updates
                {
                    State.SetStatusAndFireStateChanged(requestor, ParserState.ResolvedDeclarations);
                }

                State.SetStatusAndFireStateChanged(requestor, State.Status);
                //return; // returning here leaves state in 'ResolvedDeclarations' when a module is removed, which disables refresh
            }

            token.ThrowIfCancellationRequested();

            ExecuteCommonParseActivities(toParse.AsReadOnly(), token);
        }

        /// <summary>
        /// Clears state cache of removed components.
        /// Returns whether components have been removed.
        /// </summary>
        private bool CleanUpRemovedComponents(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token)
        {
            var removedModules = RemovedModules(modules);
            var componentRemoved = removedModules.Any();
            if (componentRemoved)
            {
                _referenceRemover.RemoveReferencesBy(removedModules, token);
                foreach (var module in removedModules)
                {
                    _moduleToModuleReferenceManager.ClearModuleToModuleReferencesFromModule(module);
                    _moduleToModuleReferenceManager.ClearModuleToModuleReferencesToModule(module);
                    State.ClearStateCache(module);
                }
            }
            return componentRemoved;
        }

        private IReadOnlyCollection<QualifiedModuleName> RemovedModules(IReadOnlyCollection<QualifiedModuleName> modules)
        {
            var modulesWithModuleDeclarations = State.DeclarationFinder.UserDeclarations(DeclarationType.Module).Select(declaration => declaration.QualifiedName.QualifiedModuleName);
            var currentlyExistingModules = modules.ToHashSet();
            var removedModuledecalrations = modulesWithModuleDeclarations.Where(module => !currentlyExistingModules.Contains(module));
            return removedModuledecalrations.ToHashSet().AsReadOnly();
        }


        public void Dispose()
        {
            State.ParseRequest -= ReparseRequested;
            Cancel(false);
        }
    }
}