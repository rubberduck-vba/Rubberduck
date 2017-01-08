using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Preprocessing;
using System.Diagnostics;
using System.IO;
using System.Linq;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Runtime.InteropServices;
// ReSharper disable LoopCanBeConvertedToQuery

namespace Rubberduck.Parsing.VBA
{
    public class ParseCoordinator : IParseCoordinator
    {
        public RubberduckParserState State { get { return _state; } }
        
        private readonly ConcurrentDictionary<IVBComponent, Tuple<Task, CancellationTokenSource>> _currentTasks =
            new ConcurrentDictionary<IVBComponent, Tuple<Task, CancellationTokenSource>>();

        private readonly IDictionary<IVBComponent, IDictionary<Tuple<string, DeclarationType>, Attributes>> _componentAttributes
            = new Dictionary<IVBComponent, IDictionary<Tuple<string, DeclarationType>, Attributes>>();

        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IAttributeParser _attributeParser;
        private readonly Func<IVBAPreprocessor> _preprocessorFactory;
        private readonly IEnumerable<ICustomDeclarationLoader> _customDeclarationLoaders;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly bool _isTestScope;

        public ParseCoordinator(
            IVBE vbe,
            RubberduckParserState state,
            IAttributeParser attributeParser,
            Func<IVBAPreprocessor> preprocessorFactory,
            IEnumerable<ICustomDeclarationLoader> customDeclarationLoaders,
            bool isTestScope = false)
        {
            _vbe = vbe;
            _state = state;
            _attributeParser = attributeParser;
            _preprocessorFactory = preprocessorFactory;
            _customDeclarationLoaders = customDeclarationLoaders;
            _isTestScope = isTestScope;

            state.ParseRequest += ReparseRequested;
        }

        // Do not access this from anywhere but ReparseRequested.
        // ReparseRequested needs to have a reference to all the cancellation tokens,
        // but the cancelees need to use their own token.
        private readonly List<CancellationTokenSource> _cancellationTokens = new List<CancellationTokenSource> {new CancellationTokenSource()};

        private void ReparseRequested(object sender, EventArgs e)
        {
            if (!_isTestScope)
            {
                Cancel();
                Task.Run(() => ParseAll(sender, _cancellationTokens[0]));
            }
            else
            {
                Parse(_cancellationTokens[0]);
            }
        }

        /// <summary>
        /// For the use of tests only
        /// </summary>
        public void Parse(CancellationTokenSource token)
        {
            State.RefreshProjects(_vbe);

            var components = new List<IVBComponent>();
            foreach (var project in State.Projects)
            {
                foreach (var component in project.VBComponents)
                {
                    components.Add(component);
                }
            }

            // tests do not fire events when components are removed--clear components
            foreach (var tree in State.ParseTrees)
            {
                State.ClearStateCache(tree.Key);    // handle potentially removed components without crashing
            }

            SyncComReferences(State.Projects);
            AddBuiltInDeclarations();

            foreach (var component in components)
            {
                State.SetModuleState(component, ParserState.Pending);
            }

            // invalidation cleanup should go into ParseAsync?
            foreach (var key in _componentAttributes.Keys)
            {
                if (!components.Contains(key))
                {
                    _componentAttributes.Remove(key);
                }
            }

            _projectDeclarations.Clear();
            State.ClearBuiltInReferences();

            var parseTasks = new Task[components.Count];
            for (var i = 0; i < components.Count; i++)
            {
                var index = i;
                parseTasks[i] = new Task(() =>
                {
                    ParseAsync(components[index], token).Wait(token.Token);

                    if (token.IsCancellationRequested)
                    {
                        return;
                    }

                    if (State.Status == ParserState.Error) { return; }

                    var qualifiedName = new QualifiedModuleName(components[index]);

                    State.SetModuleState(components[index], ParserState.ResolvingDeclarations);
                    ResolveDeclarations(qualifiedName.Component,
                        State.ParseTrees.Find(s => s.Key == qualifiedName).Value);
                });

                parseTasks[i].Start();
            }

            Task.WaitAll(parseTasks);

            if (State.Status < ParserState.Error)
            {
                State.SetStatusAndFireStateChanged(this, ParserState.ResolvedDeclarations);
                Task.WaitAll(ResolveReferencesAsync(token.Token));
                State.RebuildSelectionCache();
            }
        }

        /// <summary>
        /// Starts parsing all components of all unprotected VBProjects associated with the VBE-Instance passed to the constructor of this parser instance.
        /// </summary>
        private void ParseAll(object requestor, CancellationTokenSource token)
        {
            State.RefreshProjects(_vbe);

            var components = new List<IVBComponent>();
            foreach (var project in State.Projects)
            {
                foreach (IVBComponent component in project.VBComponents)
                {
                    components.Add(component);
                }
            }

            var componentsRemoved = false;
            foreach (var declaration in State.AllUserDeclarations)
            {
                if (!declaration.DeclarationType.HasFlag(DeclarationType.Module))
                {
                    continue;
                }

                var componentExists = false;
                foreach (var component in components)
                {
                    if (component.Name == declaration.ComponentName &&
                        component.Collection.Parent.HelpFile == declaration.ProjectId)
                    {
                        componentExists = true;
                        break;
                    }
                }

                if (!componentExists)
                {
                    componentsRemoved = true;
                    State.ClearStateCache(declaration.QualifiedName.QualifiedModuleName);
                }
            }

            var toParse = new List<IVBComponent>();
            foreach (var component in components)
            {
                if (State.IsNewOrModified(component))
                {
                    toParse.Add(component);
                }
            }

            if (toParse.Count == 0)
            {
                if (componentsRemoved)  // trigger UI updates
                {
                    State.SetStatusAndFireStateChanged(requestor, ParserState.ResolvedDeclarations);
                }

                State.SetStatusAndFireStateChanged(requestor, State.Status);
                return;
            }
            
            foreach (var component in toParse)
            {
                State.SetModuleState(component, ParserState.Pending);
            }

            SyncComReferences(State.Projects);
            AddBuiltInDeclarations();

            // invalidation cleanup should go into ParseAsync?
            foreach (var key in _componentAttributes.Keys)
            {
                if (!components.Contains(key))
                {
                    _componentAttributes.Remove(key);
                }
            }

            if (token.IsCancellationRequested)
            {
                return;
            }

            _projectDeclarations.Clear();
            State.ClearBuiltInReferences();

            var parseTasks = new Task[toParse.Count];
            for (var i = 0; i < toParse.Count; i++)
            {
                var index = i;
                parseTasks[i] = new Task(() =>
                {
                    ParseAsync(toParse[index], token).Wait(token.Token);

                    if (token.IsCancellationRequested)
                    {
                        return;
                    }

                    if (State.Status == ParserState.Error) { return; }

                    var qualifiedName = new QualifiedModuleName(toParse[index]);

                    State.SetModuleState(toParse[index], ParserState.ResolvingDeclarations);

                    ResolveDeclarations(qualifiedName.Component,
                        State.ParseTrees.Find(s => s.Key == qualifiedName).Value);
                });

                parseTasks[i].Start();
            }

            if (token.IsCancellationRequested)
            {
                return;
            }

            Task.WaitAll(parseTasks);

            if (token.IsCancellationRequested)
            {
                return;
            }

            if (State.Status < ParserState.Error)
            {
                Debug.Assert(State.ParseTrees.Count == components.Count, string.Format("ParserState has {0} parse trees for {1} components.", State.ParseTrees.Count, components.Count));
                State.SetStatusAndFireStateChanged(requestor, ParserState.ResolvedDeclarations);
                Task.WaitAll(ResolveReferencesAsync(token.Token));
                State.RebuildSelectionCache();
            }
        }

        private Task[] ResolveReferencesAsync(CancellationToken token)
        {
            foreach (var kvp in State.ParseTrees)
            {
                State.SetModuleState(kvp.Key.Component, ParserState.ResolvingReferences);
            }

            var finder = new DeclarationFinder(State.AllDeclarations, State.AllComments, State.AllAnnotations);
            var passes = new List<ICompilationPass>
                {
                    // This pass has to come first because the type binding resolution depends on it.
                    new ProjectReferencePass(finder),
                    new TypeHierarchyPass(finder, new VBAExpressionParser()),
                    new TypeAnnotationPass(finder, new VBAExpressionParser())
                };
            passes.ForEach(p => p.Execute());

            var tasks = new Task[State.ParseTrees.Count];

            for (var index = 0; index < State.ParseTrees.Count; index++)
            {
                var kvp = State.ParseTrees[index];
                if (token.IsCancellationRequested)
                {
                    return new Task[0];
                }

                tasks[index] = Task.Run(() =>
                {
                    State.SetModuleState(kvp.Key.Component, ParserState.ResolvingReferences);

                    ResolveReferences(finder, kvp.Key.Component, kvp.Value);
                }, token)
                .ContinueWith(t =>
                {
                    var undeclared = finder.Undeclared.ToList();
                    foreach (var declaration in undeclared)
                    {
                        State.AddDeclaration(declaration);
                    }
                }, token);
            }

            return tasks;
        }

        private void AddBuiltInDeclarations()
        {
            foreach (var customDeclarationLoader in _customDeclarationLoaders)
            {
                foreach (var declaration in customDeclarationLoader.Load())
                {
                    State.AddDeclaration(declaration);
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

        private void SyncComReferences(IReadOnlyList<IVBProject> projects)
        {
            var loadTasks = new List<Task>();
            var unmapped = new List<IReference>();

            foreach (var vbProject in projects)
            {
                var projectId = QualifiedModuleName.GetProjectId(vbProject);
                var references = vbProject.References;
                {
                    // use a 'for' loop to store the order of references as a 'priority'.
                    // reference resolver needs this to know which declaration to prioritize when a global identifier exists in multiple libraries.
                    for (var priority = 1; priority <= references.Count; priority++)
                    {
                        var reference = references[priority];

                        // skip loading Rubberduck.tlb (GUID is defined in AssemblyInfo.cs)
                        if (reference.Guid == "{E07C841C-14B4-4890-83E9-8C80B06DD59D}")
                        {
                            // todo: figure out why Rubberduck.tlb *sometimes* throws
                            //continue;
                        }
                        var referencedProjectId = GetReferenceProjectId(reference, projects);

                        ReferencePriorityMap map = null;
                        foreach (var item in _projectReferences)
                        {
                            if (item.ReferencedProjectId == referencedProjectId)
                            {
                                map = map != null ? null : item;
                            }
                        }

                        if (map == null)
                        {
                            map = new ReferencePriorityMap(referencedProjectId) {{projectId, priority}};
                            _projectReferences.Add(map);
                        }
                        else
                        {
                            map[projectId] = priority;
                        }

                        if (!map.IsLoaded)
                        {
                            State.OnStatusMessageUpdate(ParserState.LoadingReference.ToString());

                            var localReference = reference;

                            loadTasks.Add(
                                Task.Run(() =>
                                {
                                    try
                                    {
                                        Logger.Trace(string.Format("Loading referenced type '{0}'.", localReference.Name));

                                        var comReflector = new ReferencedDeclarationsCollector(State, localReference);
                                        if (comReflector.SerializedVersionExists)
                                        {
                                            Logger.Trace(string.Format("Deserializing reference '{0}'.", localReference.Name));
                                            foreach (var declaration in comReflector.LoadDeclarationsFromXml())
                                            {
                                                State.AddDeclaration(declaration);
                                            }
                                        }
                                        else
                                        {
                                            Logger.Trace(string.Format("COM reflecting reference '{0}'.", localReference.Name));
                                            foreach (var declaration in comReflector.LoadDeclarationsFromLibrary())
                                            {
                                                State.AddDeclaration(declaration);
                                            }
                                        }
                                    }
                                    catch (Exception exception)
                                    {
                                        unmapped.Add(reference);
                                        Logger.Warn(string.Format("Types were not loaded from referenced type library '{0}'.", reference.Name));
                                        Logger.Error(exception);
                                    }
                                }));
                            map.IsLoaded = true;
                        }
                    }
                }
            }

            var mappedIds = new List<string>();
            foreach (var item in _projectReferences)
            {
                mappedIds.Add(item.ReferencedProjectId);
            }

            foreach (var project in projects)
            {
                var references = project.References;
                {
                    foreach (var item in references)
                    {
                        if (!mappedIds.Contains(GetReferenceProjectId(item, projects)))
                        {
                            unmapped.Add(item);
                        }
                    }
                }
            }

            Task.WaitAll(loadTasks.ToArray());

            foreach (var reference in unmapped)
            {
                UnloadComReference(reference, projects);
            }
        }

        private void UnloadComReference(IReference reference, IReadOnlyList<IVBProject> projects)
        {
            var referencedProjectId = GetReferenceProjectId(reference, projects);

            ReferencePriorityMap map = null;
            foreach (var item in _projectReferences)
            {
                if (item.ReferencedProjectId == referencedProjectId)
                {
                    map = map != null ? null : item;
                }
            }
            
            if (map == null || !map.IsLoaded)
            {
                // we're removing a reference we weren't tracking? ...this shouldn't happen.
                //Debug.Assert(false);
                return;
            }
            map.Remove(referencedProjectId);
            if (map.Count == 0)
            {
                _projectReferences.Remove(map);
                State.RemoveBuiltInDeclarations(reference);
            }
        }

        private Task ParseAsync(IVBComponent component, CancellationTokenSource token, TokenStreamRewriter rewriter = null)
        {
            State.ClearStateCache(component);

            var task = new Task(() => ParseAsyncInternal(component, token.Token, rewriter));
            _currentTasks.TryAdd(component, Tuple.Create(task, token));

            Tuple<Task, CancellationTokenSource> removedTask;
            task.ContinueWith(t => _currentTasks.TryRemove(component, out removedTask), token.Token); // default also executes on cancel
            // See http://stackoverflow.com/questions/6800705/why-is-taskscheduler-current-the-default-taskscheduler
            task.Start(TaskScheduler.Default);
            return task;
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

        private void ParseAsyncInternal(IVBComponent component, CancellationToken token, TokenStreamRewriter rewriter = null)
        {
            var preprocessor = _preprocessorFactory();
            var parser = new ComponentParseTask(component, preprocessor, _attributeParser, rewriter);
            parser.ParseFailure += (sender, e) =>
            {
                State.SetModuleState(component, ParserState.Error, e.Cause as SyntaxErrorException);
            };
            parser.ParseCompleted += (sender, e) =>
            {
                lock (State)
                    lock (component)
                    {
                        State.SetModuleAttributes(component, e.Attributes);
                        State.AddParseTree(component, e.ParseTree);
                        State.AddTokenStream(component, e.Tokens);
                        State.SetModuleComments(component, e.Comments);
                        State.SetModuleAnnotations(component, e.Annotations);

                        // This really needs to go last
                        State.SetModuleState(component, ParserState.Parsed);
                    }
            };
            State.SetModuleState(component, ParserState.Parsing);
                
            parser.Start(token);
        }

        private readonly ConcurrentDictionary<string, Declaration> _projectDeclarations = new ConcurrentDictionary<string, Declaration>();
        private void ResolveDeclarations(IVBComponent component, IParseTree tree)
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
                State.SetModuleState(component, ParserState.ResolverError);
            }
            stopwatch.Stop();
            Logger.Debug("{0}ms to resolve declarations for component {1}", stopwatch.ElapsedMilliseconds, component.Name);
        }

        private Declaration CreateProjectDeclaration(QualifiedModuleName projectQualifiedName, IVBProject project)
        {
            var qualifiedName = projectQualifiedName.QualifyMemberName(project.Name);
            var projectId = qualifiedName.QualifiedModuleName.ProjectId;
            var projectDeclaration = new ProjectDeclaration(qualifiedName, project.Name, isBuiltIn: false);

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

        private void ResolveReferences(DeclarationFinder finder, IVBComponent component, IParseTree tree)
        {
            Debug.Assert(State.GetModuleState(component) == ParserState.ResolvingReferences);
            
            var qualifiedName = new QualifiedModuleName(component);
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
                    Logger.Debug("Binding resolution done for component '{0}' in {1}ms (thread {2})", component.Name,
                        watch.ElapsedMilliseconds, Thread.CurrentThread.ManagedThreadId);

                    State.SetModuleState(component, ParserState.Ready);
                }
                catch (Exception exception)
                {
                    Logger.Error(exception, "Exception thrown resolving '{0}' (thread {1}).", component.Name, Thread.CurrentThread.ManagedThreadId);
                    State.SetModuleState(component, ParserState.ResolverError);
                }
            }
        }

        public void Dispose()
        {
            State.ParseRequest -= ReparseRequested;
            Cancel(false);
        }
    }
}
