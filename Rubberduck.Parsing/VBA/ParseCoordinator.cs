using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.Symbols.DeclarationLoaders;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Preprocessing;
using System.Diagnostics;
using System.IO;
using System.Linq;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Extensions;

// ReSharper disable LoopCanBeConvertedToQuery

namespace Rubberduck.Parsing.VBA
{
    public class ParseCoordinator : IParseCoordinator
    {
        public RubberduckParserState State { get { return _state; } }

        private const int _maxDegreeOfParserParallelism = -1;
        private const int _maxDegreeOfDeclarationResolverParallelism = -1;
        private const int _maxDegreeOfReferenceResolverParallelism = -1;
        private const int _maxDegreeOfModuleStateChangeParallelism = -1;
        private const int _maxDegreeOfReferenceRemovalParallelism = -1;
        private const int _maxReferenceLoadingConcurrency = -1;

        private readonly IDictionary<IVBComponent, IDictionary<Tuple<string, DeclarationType>, Attributes>> _componentAttributes
            = new Dictionary<IVBComponent, IDictionary<Tuple<string, DeclarationType>, Attributes>>();

        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IAttributeParser _attributeParser;
        private readonly Func<IVBAPreprocessor> _preprocessorFactory;
        private readonly IEnumerable<ICustomDeclarationLoader> _customDeclarationLoaders;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly bool _isTestScope;
        private readonly string _serializedDeclarationsPath;
        private readonly IHostApplication _hostApp;

        public ParseCoordinator(
            IVBE vbe,
            RubberduckParserState state,
            IAttributeParser attributeParser,
            Func<IVBAPreprocessor> preprocessorFactory,
            IEnumerable<ICustomDeclarationLoader> customDeclarationLoaders,
            bool isTestScope = false,
            string serializedDeclarationsPath = null)
        {
            _vbe = vbe;
            _state = state;
            _attributeParser = attributeParser;
            _preprocessorFactory = preprocessorFactory;
            _customDeclarationLoaders = customDeclarationLoaders;
            _isTestScope = isTestScope;
            _serializedDeclarationsPath = serializedDeclarationsPath
                ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck", "declarations");
            _hostApp = _vbe.HostApplication();

            state.ParseRequest += ReparseRequested;
        }

        // Do not access this from anywhere but ReparseRequested.
        // ReparseRequested needs to have a reference to all the cancellation tokens,
        // but the cancelees need to use their own token.
        private readonly List<CancellationTokenSource> _cancellationTokens = new List<CancellationTokenSource> { new CancellationTokenSource() };

        private readonly Object _cancellationSyncObject = new Object();
        private readonly Object _parsingRunSyncObject = new Object();

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
                Task.Run(() => ParseAll(sender, token));
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

            var components = State.Projects.SelectMany(project => project.VBComponents).ToList();

                token.ThrowIfCancellationRequested();

            // tests do not fire events when components are removed--clear components
            ClearComponentStateCacheForTests();

                token.ThrowIfCancellationRequested();

            // invalidation cleanup should go into ParseAsync?
            CleanUpComponentAttributes(components);

                token.ThrowIfCancellationRequested();

            ExecuteCommonParseActivities(components, token);
        }

        private void ClearComponentStateCacheForTests()
        {
            foreach (var tree in State.ParseTrees)
            {
                State.ClearStateCache(tree.Key);    // handle potentially removed components without crashing
            }
        }

        private void CleanUpComponentAttributes(ICollection<IVBComponent> components)
        {
            foreach (var key in _componentAttributes.Keys)
            {
                if (!components.Contains(key))
                {
                    _componentAttributes.Remove(key);
                }
            }
        }

        private void ExecuteCommonParseActivities(ICollection<IVBComponent> toParse, CancellationToken token)
        {
                token.ThrowIfCancellationRequested();
            
            SetModuleStates(toParse, ParserState.Pending, token);

                token.ThrowIfCancellationRequested();

            SyncComReferences(State.Projects, token);
            RefreshDeclarationFinder();

                token.ThrowIfCancellationRequested();

            AddBuiltInDeclarations();
            RefreshDeclarationFinder();

                token.ThrowIfCancellationRequested();

            var modulesToParse = toParse.Select(component => new QualifiedModuleName(component)).ToHashSet();
            var toResolveReferences = ModulesForWhichToResolveReferences(modulesToParse);
            PerformPreParseCleanup(modulesToParse, toResolveReferences, token);

            ParseComponents(toParse, token);

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

            State.RebuildSelectionCache();
        }

        private void RefreshDeclarationFinder()
        {
            State.RefreshFinder(_hostApp);
        }

        private void SetModuleStates(ICollection<IVBComponent> components, ParserState parserState, CancellationToken token)
        {
            var options = new ParallelOptions();
            options.CancellationToken = token;
            options.MaxDegreeOfParallelism = _maxDegreeOfModuleStateChangeParallelism;

            Parallel.ForEach(components, options, component => State.SetModuleState(component, parserState, token, null, false));

                if (!token.IsCancellationRequested)
                {
                    State.EvaluateParserState();
                }
        }

        private ICollection<QualifiedModuleName> ModulesForWhichToResolveReferences(ICollection<QualifiedModuleName> modulesToParse)
        {
            var toResolveReferences = modulesToParse.ToHashSet();
            foreach (var qmn in modulesToParse)
            { 
                toResolveReferences.UnionWith(State.ModulesReferencing(qmn));
            }
            return toResolveReferences;
        }

        private void PerformPreParseCleanup(ICollection<QualifiedModuleName> modulesToParse, ICollection<QualifiedModuleName> toResolveReferences, CancellationToken token)
        {
            ClearModuleToModuleReferences(modulesToParse);
            RemoveAllReferencesBy(toResolveReferences, modulesToParse, State.DeclarationFinder, token); //All declarations on the modulesToParse get destroyed anyway. 
            _projectDeclarations.Clear();
        }

        private void ClearModuleToModuleReferences(ICollection<QualifiedModuleName> toClear)
        {
            foreach (var qmn in toClear)
            {
                State.ClearModuleToModuleReferencesFromModule(qmn);       
            }
        }

        //This doesn not live on the RubberduckParserState to keep concurrency haanlding out of it.
        public void RemoveAllReferencesBy(ICollection<QualifiedModuleName> referencesFromToRemove, ICollection<QualifiedModuleName> modulesNotNeedingReferenceRemoval, DeclarationFinder finder, CancellationToken token)
        {
            var referencedModulesNeedingReferenceRemoval = State.ModulesReferencedBy(referencesFromToRemove).Where(qmn => !modulesNotNeedingReferenceRemoval.Contains(qmn));

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

        private void ParseComponents(ICollection<IVBComponent> components, CancellationToken token)
        {
                token.ThrowIfCancellationRequested();
            
            SetModuleStates(components, ParserState.Parsing, token);

                token.ThrowIfCancellationRequested();

            var options = new ParallelOptions();
            options.CancellationToken = token;
            options.MaxDegreeOfParallelism = _maxDegreeOfParserParallelism;

            try
            {
                Parallel.ForEach(components,
                    options,
                    component =>
                    {
                        State.ClearStateCache(component);
                        var finishedParseTask = FinishedParseComponentTask(component, token);
                        ProcessComponentParseResults(component, finishedParseTask, token);
                    }
                );
            }
            catch (AggregateException exception)
            {
                if (exception.Flatten().InnerExceptions.All(ex => ex is OperationCanceledException))
                {
                    throw exception.InnerException; //This eliminates the stack trace, but for the cancellation, this is irrelevant.
                }
                State.SetStatusAndFireStateChanged(this, ParserState.Error);
                throw;
            }

            State.EvaluateParserState();
        }

        private Task<ComponentParseTask.ParseCompletionArgs> FinishedParseComponentTask(IVBComponent component, CancellationToken token, TokenStreamRewriter rewriter = null)
        {
            var tcs = new TaskCompletionSource<ComponentParseTask.ParseCompletionArgs>();

            var preprocessor = _preprocessorFactory();
            var parser = new ComponentParseTask(component, preprocessor, _attributeParser, rewriter);

            parser.ParseFailure += (sender, e) =>
            {
                if (e.Cause is OperationCanceledException)
                {
                    tcs.SetCanceled();
                }
                else
                {
                    tcs.SetException(e.Cause);
                }
            };
            parser.ParseCompleted += (sender, e) =>
            {
                tcs.SetResult(e);
            };

            parser.Start(token);

            return tcs.Task;
        }

        private void ProcessComponentParseResults(IVBComponent component, Task<ComponentParseTask.ParseCompletionArgs> finishedParseTask, CancellationToken token)
        {
            if (finishedParseTask.IsFaulted)
            {
                //In contrast to the situation in the success scenario, the overall parser state is reevaluated immediately.
                State.SetModuleState(component, ParserState.Error, token, finishedParseTask.Exception.InnerException as SyntaxErrorException);
            }
            else if (finishedParseTask.IsCompleted)
            {
                var result = finishedParseTask.Result;
                lock (State)
                {
                    lock (component)    
                    {
                        State.SetModuleAttributes(component, result.Attributes);
                        State.AddParseTree(component, result.ParseTree);
                        State.AddTokenStream(component, result.Tokens);
                        State.SetModuleComments(component, result.Comments);
                        State.SetModuleAnnotations(component, result.Annotations);

                        // This really needs to go last
                        //It does not reevaluate the overall parer state to avoid concurrent evaluation of all module states and for performance reasons.
                        //The evaluation has to be triggered manually in the calling procedure.
                        State.SetModuleState(component, ParserState.Parsed, token, null, false); //Note that this is ok because locks allow re-entrancy.
                    }
                }
            }
        }


        private void ResolveAllDeclarations(ICollection<IVBComponent> components, CancellationToken token)
        {
                token.ThrowIfCancellationRequested();
            
            SetModuleStates(components, ParserState.ResolvingDeclarations, token);

                token.ThrowIfCancellationRequested();

            var options = new ParallelOptions();
            options.CancellationToken = token;
            options.MaxDegreeOfParallelism = _maxDegreeOfDeclarationResolverParallelism;
            try
            {
                Parallel.ForEach(components,
                    options,
                    component =>
                    {
                        var qualifiedName = new QualifiedModuleName(component);
                        ResolveDeclarations(qualifiedName.Component,
                            State.ParseTrees.Find(s => s.Key == qualifiedName).Value, 
                            token);
                    }
                );
            }
            catch (AggregateException exception)
            {
                if (exception.Flatten().InnerExceptions.All(ex => ex is OperationCanceledException))
                {
                    throw exception.InnerException; //This eliminates the stack trace, but for the cancellation, this is irrelevant.
                }
                State.SetStatusAndFireStateChanged(this, ParserState.ResolverError);
                throw;
            }
        }

        private readonly ConcurrentDictionary<string, Declaration> _projectDeclarations = new ConcurrentDictionary<string, Declaration>();
        private void ResolveDeclarations(IVBComponent component, IParseTree tree, CancellationToken token)
        {
            if (component == null) { return; }

            var qualifiedModuleName = new QualifiedModuleName(component);

            var stopwatch = Stopwatch.StartNew();
            try
            {
                var project = component.Collection.Parent;
                var projectQualifiedName = new QualifiedModuleName(project);
                Declaration projectDeclaration;
                if (!_projectDeclarations.TryGetValue(projectQualifiedName.ProjectId, out projectDeclaration))
                {
                    projectDeclaration = CreateProjectDeclaration(projectQualifiedName, project);
                    _projectDeclarations.AddOrUpdate(projectQualifiedName.ProjectId, projectDeclaration, (s, c) => projectDeclaration);
                    State.AddDeclaration(projectDeclaration);
                }
                Logger.Debug("Creating declarations for module {0}.", qualifiedModuleName.Name);

                var declarationsListener = new DeclarationSymbolsListener(State, qualifiedModuleName, component.Type, State.GetModuleAnnotations(component), State.GetModuleAttributes(component), projectDeclaration);
                ParseTreeWalker.Default.Walk(declarationsListener, tree);
                foreach (var createdDeclaration in declarationsListener.CreatedDeclarations)
                {
                    State.AddDeclaration(createdDeclaration);
                }
            }
            catch (Exception exception)
            {
                Logger.Error(exception, "Exception thrown acquiring declarations for '{0}' (thread {1}).", component.Name, Thread.CurrentThread.ManagedThreadId);
                State.SetModuleState(component, ParserState.ResolverError, token);
            }
            stopwatch.Stop();
            Logger.Debug("{0}ms to resolve declarations for component {1}", stopwatch.ElapsedMilliseconds, component.Name);
        }

        private Declaration CreateProjectDeclaration(QualifiedModuleName projectQualifiedName, IVBProject project)
        {
            var qualifiedName = projectQualifiedName.QualifyMemberName(project.Name);
            var projectId = qualifiedName.QualifiedModuleName.ProjectId;
            var projectDeclaration = new ProjectDeclaration(qualifiedName, project.Name, false, project);

            var references = new List<ReferencePriorityMap>();
            foreach (var item in _projectReferences)
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
    
            var components = toResolve.Select(qmn => qmn.Component).ToList();
            
                token.ThrowIfCancellationRequested();
            
            SetModuleStates(components, ParserState.ResolvingReferences, token);

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
                    throw exception.InnerException; //This eliminates the stack trace, but for the cancellation, this is irrelevant.
                }
                State.SetStatusAndFireStateChanged(this, ParserState.ResolverError);
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
            var options = new ParallelOptions();
            options.CancellationToken = token;
            options.MaxDegreeOfParallelism = _maxDegreeOfReferenceResolverParallelism;
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
                    throw exception.InnerException; //This eliminates the stack trace, but for the cancellation, this is irrelevant.
                }
                State.SetStatusAndFireStateChanged(this, ParserState.ResolverError);
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
                State.AddModuleToModuleReference(referencedModule, referencingModule);
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

            State.RefreshProjects(_vbe);

                token.ThrowIfCancellationRequested();

            var components = State.Projects.SelectMany(project => project.VBComponents).ToList();

                token.ThrowIfCancellationRequested();

            var componentsRemoved = CleanUpRemovedComponents(components, token);

                token.ThrowIfCancellationRequested();

            // invalidation cleanup should go into ParseAsync?
            CleanUpComponentAttributes(components);

                token.ThrowIfCancellationRequested();

            var toParse = components.Where(component => State.IsNewOrModified(component)).ToHashSet();

                token.ThrowIfCancellationRequested();

            toParse.UnionWith(components.Where(component => State.GetModuleState(component) != ParserState.Ready));

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
        private bool CleanUpRemovedComponents(ICollection<IVBComponent> components, CancellationToken token)
        {
            var removedModuledecalrations = RemovedModuleDeclarations(components);
            var componentRemoved = removedModuledecalrations.Any();
            var removedModules = removedModuledecalrations.Select(declaration => declaration.QualifiedName.QualifiedModuleName).ToHashSet();
            if (removedModules.Any())
            {
                RemoveAllReferencesBy(removedModules, removedModules, State.DeclarationFinder, token);
                foreach (var qmn in removedModules)
                {
                    State.ClearModuleToModuleReferencesFromModule(qmn);
                    State.ClearStateCache(qmn);
                }
            }
            return componentRemoved;
        }

        private IEnumerable<Declaration> RemovedModuleDeclarations(ICollection<IVBComponent> components)
        {
            var moduleDeclarations = State.AllUserDeclarations.Where(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module));
            var componentKeys = components.Select(component => new { name = component.Name, projectId = component.Collection.Parent.HelpFile }).ToHashSet();
            var removedModuledecalrations = moduleDeclarations.Where(declaration => !componentKeys.Contains(new { name = declaration.ComponentName, projectId = declaration.ProjectId }));
            return removedModuledecalrations;
        }


        private void AddBuiltInDeclarations()
        {
            foreach (var customDeclarationLoader in _customDeclarationLoaders)
            {
                try
                {
                    foreach (var declaration in customDeclarationLoader.Load())
                    {
                        State.AddDeclaration(declaration);
                    }
                }
                catch (Exception exception)
                {
                    Logger.Error(exception, "Exception thrown adding built-in declarations. (thread {0}).", Thread.CurrentThread.ManagedThreadId);
                }
            }
        }

        private readonly HashSet<ReferencePriorityMap> _projectReferences = new HashSet<ReferencePriorityMap>();

        private string GetReferenceProjectId(IReference reference, IReadOnlyList<IVBProject> projects)
        {
            IVBProject project = null;
            foreach (var item in projects)
            {
                try
                {
                    // check the name not just the path, because path is empty in tests:
                    if (item.Name == reference.Name && item.FileName == reference.FullPath)
                    {
                        project = item;
                        break;
                    }
                }
                catch (IOException)
                {
                    // Filename throws exception if unsaved.
                }
                catch (COMException e)
                {
                    Logger.Warn(e);
                }
            }

            if (project != null)
            {
                if (string.IsNullOrEmpty(project.ProjectId))
                {
                    project.AssignProjectId();
                }
                return project.ProjectId;
            }
            return QualifiedModuleName.GetProjectId(reference);
        }

        private void SyncComReferences(IReadOnlyList<IVBProject> projects, CancellationToken token)
        {
            var unmapped = new ConcurrentBag<IReference>();

            var referencesToLoad = GetReferencesToLoadAndSaveReferencePriority(projects);
                            
            State.OnStatusMessageUpdate(ParserState.LoadingReference.ToString());

            var referenceLoadingTaskScheduler = ThrottelingTaskScheduler(_maxReferenceLoadingConcurrency); 

            //Parallel.ForEach is not used because loading the references can contain IO-bound operations.
            var loadTasks = new List<Task>();
            foreach(var reference in referencesToLoad)
            {
                var localReference = reference;
                loadTasks.Add(Task.Factory.StartNew(
                                    () => LoadReference(localReference, unmapped), 
                                    token, 
                                    TaskCreationOptions.None, 
                                    referenceLoadingTaskScheduler
                                ));
            }

            var notMappedReferences = NonMappedReferences(projects);
            foreach (var item in notMappedReferences)
            {
                unmapped.Add(item);
            }

            try
            {
                Task.WaitAll(loadTasks.ToArray(), token);
            }
            catch (AggregateException exception)
            {
                if (exception.Flatten().InnerExceptions.All(ex => ex is OperationCanceledException))
                {
                    throw exception.InnerException; //This eliminates the stack trace, but for the cancellation, this is irrelevant.
                }
                State.SetStatusAndFireStateChanged(this, ParserState.Error);
                throw;
            }
            token.ThrowIfCancellationRequested();

            foreach (var reference in unmapped)
            {
                UnloadComReference(reference, projects);
            }
        }

        private List<IReference> GetReferencesToLoadAndSaveReferencePriority(IReadOnlyList<IVBProject> projects)
        {
            var referencesToLoad = new List<IReference>();

            foreach (var vbProject in projects)
            {
                var projectId = QualifiedModuleName.GetProjectId(vbProject);
                var references = vbProject.References;

                // use a 'for' loop to store the order of references as a 'priority'.
                // reference resolver needs this to know which declaration to prioritize when a global identifier exists in multiple libraries.
                for (var priority = 1; priority <= references.Count; priority++)
                {
                    var reference = references[priority];
                    if (reference.IsBroken)
                    {
                        continue;
                    }

                    // skip loading Rubberduck.tlb (GUID is defined in AssemblyInfo.cs)
                    if (reference.Guid == "{E07C841C-14B4-4890-83E9-8C80B06DD59D}")
                    {
                        // todo: figure out why Rubberduck.tlb *sometimes* throws
                        //continue;
                    }
                    var referencedProjectId = GetReferenceProjectId(reference, projects);

                    var map = _projectReferences.FirstOrDefault(item => item.ReferencedProjectId == referencedProjectId);

                    if (map == null)
                    {
                        map = new ReferencePriorityMap(referencedProjectId) { { projectId, priority } };
                        _projectReferences.Add(map);
                    }
                    else
                    {
                        map[projectId] = priority;
                    }

                    if (!map.IsLoaded)
                    {
                        referencesToLoad.Add(reference);
                        map.IsLoaded = true;
                    }
                }
            }
            return referencesToLoad;
        }

        private TaskScheduler ThrottelingTaskScheduler(int maxLevelOfConcurrency)
        {
            if (maxLevelOfConcurrency <= 0)
            {
                return TaskScheduler.Default;
            }
            else
            {
                var taskSchedulerPair = new ConcurrentExclusiveSchedulerPair(TaskScheduler.Default, maxLevelOfConcurrency);
                return taskSchedulerPair.ConcurrentScheduler;
            }
        }

        private void LoadReference(IReference localReference, ConcurrentBag<IReference> unmapped)
        {
            Logger.Trace(string.Format("Loading referenced type '{0}'.", localReference.Name));
            var comReflector = new ReferencedDeclarationsCollector(State, localReference, _serializedDeclarationsPath);
            try
            {
                if (comReflector.SerializedVersionExists)
                {
                    LoadReferenceByDeserialization(localReference, comReflector);
                }
                else
                {
                    LoadReferenceByCOMReflection(localReference, comReflector);
                }
            }
            catch (Exception exception)
            {
                unmapped.Add(localReference);
                Logger.Warn(string.Format("Types were not loaded from referenced type library '{0}'.", localReference.Name));
                Logger.Error(exception);
            }
        }

        private void LoadReferenceByDeserialization(IReference localReference, ReferencedDeclarationsCollector comReflector)
        {
            Logger.Trace(string.Format("Deserializing reference '{0}'.", localReference.Name));
            var declarations = comReflector.LoadDeclarationsFromXml();
            foreach (var declaration in declarations)
            {
                State.AddDeclaration(declaration);
            }
        }

        private void LoadReferenceByCOMReflection(IReference localReference, ReferencedDeclarationsCollector comReflector)
        {
            Logger.Trace(string.Format("COM reflecting reference '{0}'.", localReference.Name));
            var declarations = comReflector.LoadDeclarationsFromLibrary();
            foreach (var declaration in declarations)
            {
                State.AddDeclaration(declaration);
            }
        }
        
        private List<IReference> NonMappedReferences(IReadOnlyList<IVBProject> projects)
        {
            var mappedIds = _projectReferences.Select(item => item.ReferencedProjectId).ToHashSet();
            var references = projects.SelectMany(project => project.References);
            return references.Where(item => !mappedIds.Contains(GetReferenceProjectId(item, projects))).ToList();
        }

        private void UnloadComReference(IReference reference, IReadOnlyList<IVBProject> projects)
        {
            var referencedProjectId = GetReferenceProjectId(reference, projects);

            ReferencePriorityMap map = null;
            try
            {
                map = _projectReferences.SingleOrDefault(item => item.ReferencedProjectId == referencedProjectId);
            }
            catch (InvalidOperationException exception)
            {
                //There are multiple maps with the same referencedProjectId. That should not happen. (ghost?).
                Logger.Error(exception, "Failed To unload com reference with referencedProjectID {0} because RD stores multiple instances of it.", referencedProjectId);
                return;
            }

            if (map == null || !map.IsLoaded)
            {
                // we're removing a reference we weren't tracking? ...this shouldn't happen.
                return;
            }

            map.Remove(referencedProjectId);
            if (map.Count == 0)
            {
                _projectReferences.Remove(map);
                State.RemoveBuiltInDeclarations(reference);
            }
        }


        public void Dispose()
        {
            State.ParseRequest -= ReparseRequested;
            Cancel(false);
        }
    }
}