﻿using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System.Diagnostics;
using System.Linq;
using NLog;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA
{
    /// <remarks>
    /// Note that for unit tests, TestParseCoodrinator is used in its place
    /// to support synchronous parse from BeginParse.
    /// </remarks>
    public class ParseCoordinator : IParseCoordinator
    {
        public RubberduckParserState State { get; }

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IParsingStageService _parsingStageService;
        private readonly IProjectManager _projectManager;
        private readonly IParsingCacheService _parsingCacheService;
        private readonly IParserStateManager _parserStateManager;
        private readonly IRewritingManager _rewritingManager;
        private readonly ConcurrentStack<object> _requestorStack;
        private bool _isSuspended;

        public ParseCoordinator(
            RubberduckParserState state,
            IParsingStageService parsingStageService,
            IParsingCacheService parsingCacheService,
            IProjectManager projectManager,
            IParserStateManager parserStateManager,
            IRewritingManager rewritingManager = null)
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
            _rewritingManager = rewritingManager;

            state.ParseRequest += ReparseRequested;
            state.SuspendRequest += SuspendRequested;

            _requestorStack = new ConcurrentStack<object>();
        }

        // In production, the cancellation token should be accessed inside the CancellationSyncObject
        // lock. It should not be accessible by any other code. In unit tests, however, it may be 
        // accessible in order to support synchronous parse/cancellation taking the token provided from
        // outside the parse coordinator. 
        protected CancellationTokenSource CurrentCancellationTokenSource = new CancellationTokenSource();

        protected readonly object CancellationSyncObject = new object();
        protected readonly object ParsingRunSyncObject = new object();
        protected readonly object SuspendStackSyncObject = new object();
        protected readonly ReaderWriterLockSlim ParsingSuspendLock = new ReaderWriterLockSlim(LockRecursionPolicy.NoRecursion);

        private void ReparseRequested(object sender, EventArgs e)
        {
            lock (SuspendStackSyncObject)
            {
                if (_isSuspended)
                {
                    _requestorStack.Push(sender);
                    return;
                }
            }

            BeginParse(sender);
        }

        public void SuspendRequested(object sender, RubberduckStatusSuspendParserEventArgs e)
        {
            if (ParsingSuspendLock.IsReadLockHeld)
            {
                e.Result = SuspensionResult.UnexpectedError;
                const string errorMessage =
                    "A suspension action was attempted while a read lock was held. This indicates a bug in the code logic as suspension should not be requested from same thread that has a read lock.";
                Logger.Error(errorMessage);
#if DEBUG
                Debug.Assert(false, errorMessage);
#endif
                return;
            }

            object parseRequestor = null;
            try
            {
                if (!ParsingSuspendLock.TryEnterWriteLock(e.MillisecondsTimeout))
                {
                    e.Result = SuspensionResult.TimedOut;
                    return;
                }

                lock (SuspendStackSyncObject)
                {
                    _isSuspended = true;
                }

                var originalStatus = State.Status;
                if (!e.AllowedRunStates.Contains(originalStatus))
                {
                    e.Result = SuspensionResult.IncompatibleState;
                    return;
                }
                _parserStateManager.SetStatusAndFireStateChanged(e.Requestor, ParserState.Busy,
                    CancellationToken.None);
                e.BusyAction.Invoke();
            }
            catch
            {
                e.Result = SuspensionResult.UnexpectedError;
                throw;
            }
            finally
            {
                lock (SuspendStackSyncObject)
                {
                    _isSuspended = false;
                    if (_requestorStack.TryPop(out var lastRequestor))
                    {
                        _requestorStack.Clear();
                        parseRequestor = lastRequestor;
                    }

                    // Though there were no reparse requests, we need to reset the state before we release the 
                    // write lock to avoid introducing discrepancy in the parser state due to readers being 
                    // blocked. Any reparse requests must be done outside the write lock; see further below.
                    if (parseRequestor == null)
                    {
                        // We cannot make any assumptions about the original state, nor do we know
                        // anything about resuming the previous state, so we must delegate the state
                        // evaluation to the state manager.
                        _parserStateManager.EvaluateOverallParserState(CancellationToken.None);
                    }
                }

                if (ParsingSuspendLock.IsWriteLockHeld)
                {
                    ParsingSuspendLock.ExitWriteLock();
                }

                if (e.Result == SuspensionResult.Pending)
                {
                    e.Result = SuspensionResult.Completed;
                }
            }

            // Any reparse requests must be done outside the write lock to avoid deadlocks
            if (parseRequestor != null)
            {
                BeginParse(parseRequestor);
            }
        }

        /// <remarks>
        /// Overriden in the unit test project to facilicate synchronous unit tests
        /// Refer to TestParserCoordinator class in the unit test project.
        /// </remarks>
        public virtual void BeginParse(object sender)
        {
            Task.Run(() => ParseAll(sender));
        }

        private void Cancel(bool createNewTokenSource = true)
        {
            lock (CancellationSyncObject)
            {
                if (CurrentCancellationTokenSource == null)
                {
                    Logger.Error("Tried to cancel a parse after the final cancellation.");
                    return;
                }

                var oldTokenSource = CurrentCancellationTokenSource;
                CurrentCancellationTokenSource = createNewTokenSource ? new CancellationTokenSource() : null;

                oldTokenSource.Cancel();
                oldTokenSource.Dispose();
            }
        }

        private void ExecuteCommonParseActivities(IReadOnlyCollection<QualifiedModuleName> toParse, IReadOnlyCollection<QualifiedModuleName> toReresolveReferencesInput, IReadOnlyCollection<string> newProjectIds, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            var toReresolveReferences = new HashSet<QualifiedModuleName>();
            toReresolveReferences.UnionWith(toReresolveReferencesInput);
            token.ThrowIfCancellationRequested();

            _parserStateManager.SetModuleStates(toParse, ParserState.Started, token);
            token.ThrowIfCancellationRequested();

            _parsingCacheService.ClearProjectWhoseCompilationArgumentsChanged();
            token.ThrowIfCancellationRequested();

            _parserStateManager.SetStatusAndFireStateChanged(this, ParserState.LoadingReference, token);
            token.ThrowIfCancellationRequested();

            _parsingStageService.SyncComReferences(token);
            if (_parsingStageService.LastSyncOfCOMReferencesLoadedReferences || _parsingStageService.COMReferencesUnloadedInLastSync.Any())
            {
                var unloadedReferences = _parsingStageService.COMReferencesUnloadedInLastSync.ToHashSet();
                var unloadedModules =
                    _parsingCacheService.DeclarationFinder.AllModules
                        .Where(qmn => unloadedReferences.Contains(qmn.ProjectId))
                        .ToHashSet();
                var additionalModulesToBeReresolved = OtherModulesReferencingAnyNotToBeParsed(unloadedModules.AsReadOnly(), toParse);
                toReresolveReferences.UnionWith(additionalModulesToBeReresolved);
                _parserStateManager.SetModuleStates(additionalModulesToBeReresolved, ParserState.ResolvingReferences, token);
                ClearModuleToModuleReferences(unloadedModules);
                RefreshDeclarationFinder();
            }

            if (_parsingStageService.COMReferencesAffectedByPriorityChangesInLastSync.Any())
            {
                //We only use the referencedProjectId because that simplifies the reference management immensely.  
                var affectedReferences = _parsingStageService.COMReferencesAffectedByPriorityChangesInLastSync
                    .Select(tpl => tpl.referencedProjectId)
                    .ToHashSet();
                var referenceModules =
                    _parsingCacheService.DeclarationFinder.AllModules
                        .Where(qmn => affectedReferences.Contains(qmn.ProjectId))
                        .ToHashSet();
                var additionalModulesToBeReresolved = OtherModulesReferencingAnyNotToBeParsed(referenceModules.AsReadOnly(), toParse);
                toReresolveReferences.UnionWith(additionalModulesToBeReresolved);
                _parserStateManager.SetModuleStates(additionalModulesToBeReresolved, ParserState.ResolvingReferences, token);
            }
            token.ThrowIfCancellationRequested();

            _parsingStageService.LoadBuitInDeclarations();
            if (newProjectIds.Any())
            {
                _parsingStageService.CreateProjectDeclarations(newProjectIds);
                RefreshDeclarationFinder();
            }
            if (_parsingStageService.LastLoadOfBuiltInDeclarationsLoadedDeclarations || newProjectIds.Any())
            { 
                RefreshDeclarationFinder();
            }
            token.ThrowIfCancellationRequested();

            _parsingStageService.RefreshProjectReferences();
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
            _parserStateManager.SetStatusAndFireStateChanged(this, ParserState.ResolvedDeclarations, token);

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
        private void ParseAll(object requestor)
        {
            CancellationToken token;
            Stopwatch watch = null;
            var lockTaken = false;
            try
            {
                if (!ParsingSuspendLock.IsWriteLockHeld)
                {
                    ParsingSuspendLock.EnterReadLock();
                }
                lock (CancellationSyncObject)
                {
                    Cancel();
                    token = CurrentCancellationTokenSource.Token;
                }
                Monitor.Enter(ParsingRunSyncObject, ref lockTaken);
                
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
                if (_parserStateManager.OverallParserState != ParserState.UnexpectedError)
                {
                    _parserStateManager.SetStatusAndFireStateChanged(this, ParserState.UnexpectedError, token);
                }
            }
            finally
            {
                if (watch != null && watch.IsRunning) watch.Stop();
                if (lockTaken) Monitor.Exit(ParsingRunSyncObject);
                if (ParsingSuspendLock.IsReadLockHeld)
                {
                    ParsingSuspendLock.ExitReadLock();
                }
            }
            if (watch != null) Logger.Debug("Parsing run finished after {0}s. (thread {1}).", watch.Elapsed.TotalSeconds, Thread.CurrentThread.ManagedThreadId);
        }

        protected void ParseAllInternal(object requestor, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            _parserStateManager.SetStatusAndFireStateChanged(requestor, ParserState.Started, token);
            token.ThrowIfCancellationRequested();

            _rewritingManager?.InvalidateAllSessions();
            token.ThrowIfCancellationRequested();

            _projectManager.RefreshProjects();
            token.ThrowIfCancellationRequested();

            var modules = _projectManager.AllModules();
            token.ThrowIfCancellationRequested();

            var projects = _projectManager.Projects;
            var projectIds = projects.Select(tpl => tpl.ProjectId).ToList().AsReadOnly();
            token.ThrowIfCancellationRequested();

            var toParse = modules.Where(module => State.IsNewOrModified(module)).ToHashSet();
            token.ThrowIfCancellationRequested();

            toParse.UnionWith(modules.Where(module => _parserStateManager.GetModuleState(module) != ParserState.Ready));
            token.ThrowIfCancellationRequested();

            _parsingCacheService.ReloadCompilationArguments(projectIds);
            token.ThrowIfCancellationRequested();

            var projectsWithChangedCompilationArguments = _parsingCacheService.ProjectWhoseCompilationArgumentsChanged();
            token.ThrowIfCancellationRequested();

            toParse.UnionWith(ModulesInProjects(projectsWithChangedCompilationArguments));
            token.ThrowIfCancellationRequested();

            toParse = toParse.Where(module => module.IsParsable).ToHashSet();
            token.ThrowIfCancellationRequested();

            var removedModules = RemovedModules(modules);
            token.ThrowIfCancellationRequested();

            var removedProjects = RemovedProjects(projects.Select(tpl => tpl.Project).ToList().AsReadOnly());
            token.ThrowIfCancellationRequested();

            removedModules.UnionWith(ModulesInProjects(removedProjects));
            token.ThrowIfCancellationRequested();

            var notRemovedDisposedProjects = NotRemovedDisposedProjects(removedProjects);
            toParse.UnionWith(ModulesInProjects(notRemovedDisposedProjects));

            var toReResolveReferences = OtherModulesReferencingAnyNotToBeParsed(removedModules.AsReadOnly(), toParse.AsReadOnly());
            token.ThrowIfCancellationRequested();

            _parserStateManager.SetModuleStates(toReResolveReferences, ParserState.ResolvingReferences, token);
            token.ThrowIfCancellationRequested();

            CleanUpRemovedComponents(removedModules.AsReadOnly(), token);
            token.ThrowIfCancellationRequested();

            //This must come after the component cleanup because of cache invalidation.
            CleanUpProjects(removedProjects);
            token.ThrowIfCancellationRequested();

            CleanUpProjects(notRemovedDisposedProjects);
            token.ThrowIfCancellationRequested();

            var newProjects = NewProjects(projectIds);

            ExecuteCommonParseActivities(toParse.AsReadOnly(), toReResolveReferences, newProjects, token);
        }

        private IReadOnlyCollection<string> NewProjects(IReadOnlyCollection<string> projectIds)
        {
            var existingProjects = State.DeclarationFinder
                .UserDeclarations(DeclarationType.Project)
                .Select(declaration => declaration.ProjectId)
                .ToHashSet();
            var newProjects = projectIds.Where(projectId => !existingProjects
                    .Contains(projectId))
                    .ToList()
                    .AsReadOnly();
            return newProjects;
        }

        private IReadOnlyCollection<QualifiedModuleName> OtherModulesReferencingAnyNotToBeParsed(IReadOnlyCollection<QualifiedModuleName> removedModules, IReadOnlyCollection<QualifiedModuleName> toParse)
        {
            return _parsingCacheService.ModulesReferencingAny(removedModules)
                        .Where(qmn => !removedModules.Contains(qmn) && !toParse.Contains(qmn))
                        .ToHashSet().AsReadOnly();
        }

        private IEnumerable<QualifiedModuleName> ModulesInProjects(IReadOnlyCollection<string> projectIds)
        {
            return State.DeclarationFinder.UserDeclarations(DeclarationType.Module)
                    .Where(declaration => projectIds.Contains(declaration.ProjectId))
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

        private void CleanUpProjects(IReadOnlyCollection<string> removedProjectIds)
        {
            _parsingCacheService.RemoveCompilationArgumentsFromCache(removedProjectIds);
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

        private IReadOnlyCollection<string> NotRemovedDisposedProjects(IEnumerable<string> removedProjects)
        {
            var disposedProjects = State.DeclarationFinder
                                    .UserDeclarations(DeclarationType.Project)
                                    .OfType<ProjectDeclaration>()
                                    .Where(declaration => declaration.IsDisposed)
                                    .Select(declaration => declaration.ProjectId)
                                    .ToHashSet();
            disposedProjects.ExceptWith(removedProjects);
            return disposedProjects.AsReadOnly();
        }

        public void Dispose()
        {
            State.ParseRequest -= ReparseRequested;
            Cancel(false);
            ParsingSuspendLock.Dispose();
        }
    }
}