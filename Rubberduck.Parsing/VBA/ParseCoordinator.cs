using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System.Diagnostics;
using System.Linq;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.VBEditor.Application;

// ReSharper disable LoopCanBeConvertedToQuery

namespace Rubberduck.Parsing.VBA
{
    public class ParseCoordinator : IParseCoordinator
    {
        public RubberduckParserState State { get { return _state; } }

        private const int _maxDegreeOfDeclarationResolverParallelism = -1;
        private const int _maxDegreeOfReferenceResolverParallelism = -1;
        private const int _maxDegreeOfReferenceRemovalParallelism = -1;

        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IModuleToModuleReferenceManager _moduleToModuleReferenceManager;
        private readonly IParserStateManager _parserStateManager;
        private readonly ICOMReferenceManager _comReferenceManager;
        private readonly IBuiltInDeclarationLoader _builtInDeclarationLoader;
        private readonly IParseRunner _parseRunner;

        private readonly bool _isTestScope;
        private readonly IHostApplication _hostApp;

        public ParseCoordinator(
            IVBE vbe,
            RubberduckParserState state,
            IModuleToModuleReferenceManager moduleToModuleReferenceManager,
            IParserStateManager parserStateManager,
            ICOMReferenceManager comSynchronizationRunner,
            IBuiltInDeclarationLoader builtInDeclarationLoader,
            IParseRunner parseRunner,
            bool isTestScope = false)
        {
            _vbe = vbe;
            _state = state;
            _moduleToModuleReferenceManager = moduleToModuleReferenceManager;
            _parserStateManager = parserStateManager;
            _comReferenceManager = comSynchronizationRunner;
            _builtInDeclarationLoader = builtInDeclarationLoader;
            _parseRunner = parseRunner;
            _isTestScope = isTestScope;
            _hostApp = _vbe.HostApplication();

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
            
            State.RefreshProjects(_vbe);

                token.ThrowIfCancellationRequested();

            var modules = State.Projects.SelectMany(project => project.VBComponents).Select(component => new QualifiedModuleName(component)).ToHashSet();

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

        private void ExecuteCommonParseActivities(ICollection<QualifiedModuleName> toParse, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            _parserStateManager.SetModuleStates(toParse, ParserState.Pending, token);

            token.ThrowIfCancellationRequested();

            _comReferenceManager.SyncComReferences(State.Projects, token);
            if (_comReferenceManager.LastRunLoadedReferences || _comReferenceManager.LastRunUnloadedReferences)
            {
                RefreshDeclarationFinder();
            }

            token.ThrowIfCancellationRequested();

            _builtInDeclarationLoader.LoadBuitInDeclarations();
            if (_builtInDeclarationLoader.LastExecutionLoadedDeclarations)
            { 
                RefreshDeclarationFinder();
            }

            token.ThrowIfCancellationRequested();

            var toResolveReferences = ModulesForWhichToResolveReferences(toParse);
            PerformPreParseCleanup(toParse, toResolveReferences, token);

            _parseRunner.ParseModules(toParse, token);

            if (token.IsCancellationRequested || State.Status >= ParserState.Error)
            {
                throw new OperationCanceledException(token);
            }

            ResolveAllDeclarations(toParse, token);
            RefreshDeclarationFinder();

            if (token.IsCancellationRequested || State.Status >= ParserState.Error)
            {
                throw new OperationCanceledException(token);
            }

            State.SetStatusAndFireStateChanged(this, ParserState.ResolvedDeclarations);

            if (token.IsCancellationRequested || State.Status >= ParserState.Error)
            {
                throw new OperationCanceledException(token);
            }

            ResolveAllReferences(toResolveReferences, token);

            if (token.IsCancellationRequested || State.Status >= ParserState.Error)
            {
                throw new OperationCanceledException(token);
            }
        }

        private void RefreshDeclarationFinder()
        {
            State.RefreshFinder(_hostApp);
        }

        private ICollection<QualifiedModuleName> ModulesForWhichToResolveReferences(ICollection<QualifiedModuleName> modulesToParse)
        {
            var toResolveReferences = modulesToParse.ToHashSet();
            toResolveReferences.UnionWith(_moduleToModuleReferenceManager.ModulesReferencingAny(modulesToParse));
            return toResolveReferences;
        }

        private void PerformPreParseCleanup(ICollection<QualifiedModuleName> modulesToParse, ICollection<QualifiedModuleName> toResolveReferences, CancellationToken token)
        {
            _moduleToModuleReferenceManager.ClearModuleToModuleReferencesFromModule(modulesToParse);
            RemoveAllReferencesBy(toResolveReferences, modulesToParse, State.DeclarationFinder, token); //All declarations on the modulesToParse get destroyed anyway. 
            _projectDeclarations.Clear();
        }

        //This does not live on the RubberduckParserState to keep concurrency haanlding out of it.
        public void RemoveAllReferencesBy(ICollection<QualifiedModuleName> referencesFromToRemove, ICollection<QualifiedModuleName> modulesNotNeedingReferenceRemoval, DeclarationFinder finder, CancellationToken token)
        {
            var referencedModulesNeedingReferenceRemoval = _moduleToModuleReferenceManager.ModulesReferencedByAny(referencesFromToRemove).Where(qmn => !modulesNotNeedingReferenceRemoval.Contains(qmn));

            var options = new ParallelOptions();
            options.CancellationToken = token;
            options.MaxDegreeOfParallelism = _maxDegreeOfReferenceRemovalParallelism;

            Parallel.ForEach(referencedModulesNeedingReferenceRemoval, options, qmn => RemoveReferences(finder.Members(qmn), referencesFromToRemove));
        }

        private void RemoveReferences(IEnumerable<Declaration> declarations, ICollection<QualifiedModuleName> referencesFromToRemove)
        {
            foreach (var declaration in declarations)
            {
                declaration.RemoveReferencesFrom(referencesFromToRemove);
            }
        }


        private void ResolveAllDeclarations(ICollection<QualifiedModuleName> modules, CancellationToken token)
        {
                token.ThrowIfCancellationRequested();
            
            _parserStateManager.SetModuleStates(modules, ParserState.ResolvingDeclarations, token);

                token.ThrowIfCancellationRequested();

            var options = new ParallelOptions();
            options.CancellationToken = token;
            options.MaxDegreeOfParallelism = _maxDegreeOfDeclarationResolverParallelism;
            try
            {
                Parallel.ForEach(modules,
                    options,
                    module =>
                    {
                        ResolveDeclarations(module,
                            State.ParseTrees.Find(s => s.Key == module).Value, 
                            token);
                    }
                );
            }
            catch (AggregateException exception)
            {
                if (exception.Flatten().InnerExceptions.All(ex => ex is OperationCanceledException))
                {
                    throw exception.InnerException ?? exception; //This eliminates the stack trace, but for the cancellation, this is irrelevant.
                }
                _parserStateManager.SetStatusAndFireStateChanged(this, ParserState.ResolverError, token);
                throw;
            }
        }

        private readonly ConcurrentDictionary<string, Declaration> _projectDeclarations = new ConcurrentDictionary<string, Declaration>();
        private void ResolveDeclarations(QualifiedModuleName module, IParseTree tree, CancellationToken token)
        {
            if (module == null) { return; }

            var stopwatch = Stopwatch.StartNew();
            try
            {
                var project = module.Component.Collection.Parent;
                var projectQualifiedName = new QualifiedModuleName(project);
                Declaration projectDeclaration;
                if (!_projectDeclarations.TryGetValue(projectQualifiedName.ProjectId, out projectDeclaration))
                {
                    projectDeclaration = CreateProjectDeclaration(projectQualifiedName, project);
                    _projectDeclarations.AddOrUpdate(projectQualifiedName.ProjectId, projectDeclaration, (s, c) => projectDeclaration);
                    State.AddDeclaration(projectDeclaration);
                }
                Logger.Debug("Creating declarations for module {0}.", module.Name);

                var declarationsListener = new DeclarationSymbolsListener(State, module, module.ComponentType, State.GetModuleAnnotations(module), State.GetModuleAttributes(module), projectDeclaration);
                ParseTreeWalker.Default.Walk(declarationsListener, tree);
                foreach (var createdDeclaration in declarationsListener.CreatedDeclarations)
                {
                    State.AddDeclaration(createdDeclaration);
                }
            }
            catch (Exception exception)
            {
                Logger.Error(exception, "Exception thrown acquiring declarations for '{0}' (thread {1}).", module.Name, Thread.CurrentThread.ManagedThreadId);
                _parserStateManager.SetModuleState(module, ParserState.ResolverError, token);
            }
            stopwatch.Stop();
            Logger.Debug("{0}ms to resolve declarations for component {1}", stopwatch.ElapsedMilliseconds, module.Name);
        }

        private Declaration CreateProjectDeclaration(QualifiedModuleName projectQualifiedName, IVBProject project)
        {
            var qualifiedName = projectQualifiedName.QualifyMemberName(project.Name);
            var projectId = qualifiedName.QualifiedModuleName.ProjectId;
            var projectDeclaration = new ProjectDeclaration(qualifiedName, project.Name, true, project);

            var references = new List<ReferencePriorityMap>();
            foreach (var item in _comReferenceManager.ProjectReferences)
            {
                if (item.ContainsKey(projectId))
                {
                    references.Add(item);
                }
            }

            foreach (var reference in references)
            {
                int priority = reference[projectId];
                projectDeclaration.AddProjectReference(reference.ReferencedProjectId, priority);
            }
            return projectDeclaration;
        }


        private void ResolveAllReferences(ICollection<QualifiedModuleName> toResolve, CancellationToken token)
        {
            
                token.ThrowIfCancellationRequested();
            
            _parserStateManager.SetModuleStates(toResolve, ParserState.ResolvingReferences, token);

                token.ThrowIfCancellationRequested();

            ExecuteCompilationPasses();

                token.ThrowIfCancellationRequested();

            var parseTreesToResolve = State.ParseTrees.Where(kvp => toResolve.Contains(kvp.Key)).ToList();

            var options = new ParallelOptions();
            options.CancellationToken = token;
            options.MaxDegreeOfParallelism = _maxDegreeOfReferenceResolverParallelism;
            try
            {
                Parallel.For(0, parseTreesToResolve.Count, options,
                    (index) => ResolveReferences(State.DeclarationFinder, parseTreesToResolve[index].Key, parseTreesToResolve[index].Value, token)
                );
            }
            catch (AggregateException exception)
            {
                if (exception.Flatten().InnerExceptions.All(ex => ex is OperationCanceledException))
                {
                    throw exception.InnerException ?? exception; //This eliminates the stack trace, but for the cancellation, this is irrelevant.
                }
                _parserStateManager.SetStatusAndFireStateChanged(this, ParserState.ResolverError, token);
                throw;
            }

                token.ThrowIfCancellationRequested();

            AddModuleToModuleReferences(State.DeclarationFinder, token);

                token.ThrowIfCancellationRequested();
            
            AddNewUndeclaredVariablesToDeclarations();
            AddNewUnresolvedMemberDeclarations();

            //This is here and not in the calling method because it has to happen before the ready state is reached.
            RefreshDeclarationFinder();

                token.ThrowIfCancellationRequested();

            State.EvaluateParserState();
        }

        private void ExecuteCompilationPasses()
        {
            var passes = new List<ICompilationPass>
                {
                    // This pass has to come first because the type binding resolution depends on it.
                    new ProjectReferencePass(State.DeclarationFinder),
                    new TypeHierarchyPass(State.DeclarationFinder, new VBAExpressionParser()),
                    new TypeAnnotationPass(State.DeclarationFinder, new VBAExpressionParser())
                };
            passes.ForEach(p => p.Execute());
        }

        private void ResolveReferences(DeclarationFinder finder, QualifiedModuleName qualifiedName, IParseTree tree, CancellationToken token)
        {
            Debug.Assert(State.GetModuleState(qualifiedName.Component) == ParserState.ResolvingReferences || token.IsCancellationRequested);

                token.ThrowIfCancellationRequested();

            Logger.Debug("Resolving identifier references in '{0}'... (thread {1})", qualifiedName.Name, Thread.CurrentThread.ManagedThreadId);

            var resolver = new IdentifierReferenceResolver(qualifiedName, finder);
            var listener = new IdentifierReferenceListener(resolver);

            if (!string.IsNullOrWhiteSpace(tree.GetText().Trim()))
            {
                var walker = new ParseTreeWalker();
                try
                {
                    var watch = Stopwatch.StartNew();
                    walker.Walk(listener, tree);
                    watch.Stop();
                    Logger.Debug("Binding resolution done for component '{0}' in {1}ms (thread {2})", qualifiedName.Name,
                        watch.ElapsedMilliseconds, Thread.CurrentThread.ManagedThreadId);

                    //Evaluation of the overall status has to be defered to allow processing of undeclared variables before setting the ready state.
                    State.SetModuleState(qualifiedName.Component, ParserState.Ready, token, null, false);
                }
                catch (OperationCanceledException)
                {
                    throw;  //We do not want to set an error state if the exception was just caused by some cancellation.
                }
                catch (Exception exception)
                {
                    Logger.Error(exception, "Exception thrown resolving '{0}' (thread {1}).", qualifiedName.Name, Thread.CurrentThread.ManagedThreadId);
                    State.SetModuleState(qualifiedName.Component, ParserState.ResolverError, token);
                }
            }
        }

        private void AddModuleToModuleReferences(DeclarationFinder finder, CancellationToken token)
        {
            var options = new ParallelOptions
            {
                CancellationToken = token,
                MaxDegreeOfParallelism = _maxDegreeOfReferenceResolverParallelism
            };

            try
            {
                Parallel.For(0, State.ParseTrees.Count, options,
                    (index) => AddModuleToModuleReferences(finder, State.ParseTrees[index].Key)
                );
            }
            catch (AggregateException exception)
            {
                if (exception.Flatten().InnerExceptions.All(ex => ex is OperationCanceledException))
                {
                    throw exception.InnerException ?? exception; //This eliminates the stack trace, but for the cancellation, this is irrelevant.
                }
                
                _parserStateManager.SetStatusAndFireStateChanged(this, ParserState.ResolverError, token);
                throw;
            }
        }

        private void AddModuleToModuleReferences(DeclarationFinder finder, QualifiedModuleName referencedModule)
        {
            var referencingModules = finder.Members(referencedModule)
                                        .SelectMany(declaration => declaration.References)
                                        .Select(reference => reference.QualifiedModuleName)
                                        .Distinct()
                                        .Where(referencingModule => !referencedModule.Equals(referencingModule));
            foreach (var referencingModule in referencingModules)
            {
                _moduleToModuleReferenceManager.AddModuleToModuleReference(referencingModule, referencedModule);
            }
        }

        private void AddNewUndeclaredVariablesToDeclarations()
        {
            var undeclared = State.DeclarationFinder.FreshUndeclared.ToList();
            foreach (var declaration in undeclared)
            {
                State.AddDeclaration(declaration);
            }
        }

        private void AddNewUnresolvedMemberDeclarations()
        {
            var unresolved = State.DeclarationFinder.FreshUnresolvedMemberDeclarations().ToList();
            foreach (var declaration in unresolved)
            {
                State.AddUnresolvedMemberDeclaration(declaration);
            }
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
                State.SetStatusAndFireStateChanged(this, ParserState.Error);
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

            State.RefreshProjects(_vbe);

                token.ThrowIfCancellationRequested();

            var modules = State.Projects.SelectMany(project => project.VBComponents).Select(component => new QualifiedModuleName(component)).ToList();

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

            ExecuteCommonParseActivities(toParse, token);
        }

        /// <summary>
        /// Clears state cache of removed components.
        /// Returns whether components have been removed.
        /// </summary>
        private bool CleanUpRemovedComponents(ICollection<QualifiedModuleName> modules, CancellationToken token)
        {
            var removedModules = RemovedModules(modules).ToHashSet();
            var componentRemoved = removedModules.Any();
            if (componentRemoved)
            {
                RemoveAllReferencesBy(removedModules, removedModules, State.DeclarationFinder, token);
                foreach (var module in removedModules)
                {
                    State.ClearModuleToModuleReferencesFromModule(module);
                    State.ClearStateCache(module);
                }
            }
            return componentRemoved;
        }

        private IEnumerable<QualifiedModuleName> RemovedModules(ICollection<QualifiedModuleName> modules)
        {
            var modulesWithModuleDeclarations = State.DeclarationFinder.UserDeclarations(DeclarationType.Module).Select(declaration => declaration.QualifiedName.QualifiedModuleName);
            var currentlyExistingModules = modules.ToHashSet();
            var removedModuledecalrations = modulesWithModuleDeclarations.Where(module => !currentlyExistingModules.Contains(module));
            return removedModuledecalrations;
        }


        public void Dispose()
        {
            State.ParseRequest -= ReparseRequested;
            Cancel(false);
        }
    }
}