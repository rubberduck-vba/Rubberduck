using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Preprocessing;
using System.Diagnostics;
using Rubberduck.VBEditor.Extensions;
using System.IO;
using NLog;
// ReSharper disable LoopCanBeConvertedToQuery

namespace Rubberduck.Parsing.VBA
{
    public class RubberduckParser : IRubberduckParser
    {
        public RubberduckParserState State { get { return _state; } }

        private CancellationTokenSource _central = new CancellationTokenSource();
        private CancellationTokenSource _resolverTokenSource; // linked to _central later
        private readonly ConcurrentDictionary<VBComponent, Tuple<Task, CancellationTokenSource>> _currentTasks =
            new ConcurrentDictionary<VBComponent, Tuple<Task, CancellationTokenSource>>();

        private readonly IDictionary<VBComponent, IDictionary<Tuple<string, DeclarationType>, Attributes>> _componentAttributes
            = new Dictionary<VBComponent, IDictionary<Tuple<string, DeclarationType>, Attributes>>();

        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IAttributeParser _attributeParser;
        private readonly Func<IVBAPreprocessor> _preprocessorFactory;
        private readonly IEnumerable<ICustomDeclarationLoader> _customDeclarationLoaders;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public RubberduckParser(
            VBE vbe,
            RubberduckParserState state,
            IAttributeParser attributeParser,
            Func<IVBAPreprocessor> preprocessorFactory,
            IEnumerable<ICustomDeclarationLoader> customDeclarationLoaders)
        {
            _resolverTokenSource = CancellationTokenSource.CreateLinkedTokenSource(_central.Token);
            _vbe = vbe;
            _state = state;
            _attributeParser = attributeParser;
            _preprocessorFactory = preprocessorFactory;
            _customDeclarationLoaders = customDeclarationLoaders;

            state.ParseRequest += ReparseRequested;
        }

        private void ReparseRequested(object sender, ParseRequestEventArgs e)
        {
            if (e.IsFullReparseRequest)
            {
                Cancel();
                Task.Run(() => ParseAll());
            }
            else
            {
                Cancel(e.Component);
                Task.Run(() =>
                {
                    SyncComReferences(State.Projects);
                    AddBuiltInDeclarations();

                    if (_resolverTokenSource.IsCancellationRequested || _central.IsCancellationRequested)
                    {
                        return;
                    }

                    ParseAsync(e.Component, CancellationToken.None).Wait();

                    if (_resolverTokenSource.IsCancellationRequested || _central.IsCancellationRequested)
                    {
                        return;
                    }

                    if (State.Status == ParserState.Error) { return; }

                    var qualifiedName = new QualifiedModuleName(e.Component);
                    Logger.Debug("Module '{0}' {1}", qualifiedName.ComponentName,
                        State.IsNewOrModified(qualifiedName) ? "was modified" : "was NOT modified");

                    State.SetModuleState(e.Component, ParserState.ResolvingDeclarations);
                    ResolveDeclarations(qualifiedName.Component,
                        State.ParseTrees.Find(s => s.Key == qualifiedName).Value);
                    
                    if (State.Status < ParserState.Error)
                    {
                        State.SetStatusAndFireStateChanged(ParserState.ResolvedDeclarations);
                        ResolveReferencesAsync();
                    }
                });
            }
        }

        /// <summary>
        /// For the use of tests only
        /// </summary>
        public void Parse()
        {
            if (State.Projects.Count == 0)
            {
                foreach (var project in _vbe.VBProjects.UnprotectedProjects())
                {
                    State.AddProject(project);
                }
            }

            var components = new List<VBComponent>();
            foreach (var project in State.Projects)
            {
                foreach (VBComponent component in project.VBComponents)
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
                    ParseAsync(components[index], CancellationToken.None).Wait();

                    if (_resolverTokenSource.IsCancellationRequested || _central.IsCancellationRequested)
                    {
                        return;
                    }

                    if (State.Status == ParserState.Error) { return; }

                    var qualifiedName = new QualifiedModuleName(components[index]);
                    Logger.Debug("Module '{0}' {1}", qualifiedName.ComponentName,
                        State.IsNewOrModified(qualifiedName) ? "was modified" : "was NOT modified");

                    State.SetModuleState(components[index], ParserState.ResolvingDeclarations);
                    ResolveDeclarations(qualifiedName.Component,
                        State.ParseTrees.Find(s => s.Key == qualifiedName).Value);
                });

                parseTasks[i].Start();
            }

            Task.WaitAll(parseTasks);

            if (State.Status < ParserState.Error)
            {
                State.SetStatusAndFireStateChanged(ParserState.ResolvedDeclarations);
                Task.WaitAll(ResolveReferencesAsync());
            }
        }

        /// <summary>
        /// Starts parsing all components of all unprotected VBProjects associated with the VBE-Instance passed to the constructor of this parser instance.
        /// </summary>
        private void ParseAll()
        {
            if (State.Projects.Count == 0)
            {
                foreach (var project in _vbe.VBProjects.UnprotectedProjects())
                {
                    State.AddProject(project);
                }
            }

            var components = new List<VBComponent>();
            foreach (var project in State.Projects)
            {
                foreach (VBComponent component in project.VBComponents)
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

            var toParse = new List<VBComponent>();
            var unchanged = new List<VBComponent>();

            foreach (var component in components)
            {
                if (State.IsNewOrModified(component))
                {
                    toParse.Add(component);
                }
                else
                {
                    unchanged.Add(component);
                }
            }

            SyncComReferences(State.Projects);
            AddBuiltInDeclarations();

            if (toParse.Count == 0)
            {
                if (componentsRemoved)  // trigger UI updates
                {
                    State.SetStatusAndFireStateChanged(ParserState.ResolvedDeclarations);
                }

                State.SetStatusAndFireStateChanged(State.Status);
                return;
            }
            
            lock (State)  // note, method is invoked from UI thread... really need the lock here?
            {
                foreach (var component in toParse)
                {
                    State.SetModuleState(component, ParserState.Pending);
                }
                foreach (var component in unchanged)
                {
                    // note: seting to 'Parsed' would include them in the resolver walk. 'Ready' excludes them.
                    State.SetModuleState(component, ParserState.Ready);
                }
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

            var parseTasks = new Task[toParse.Count];
            for (var i = 0; i < toParse.Count; i++)
            {
                var index = i;
                parseTasks[i] = new Task(() =>
                {
                    ParseAsync(toParse[index], CancellationToken.None).Wait();

                    if (_resolverTokenSource.IsCancellationRequested || _central.IsCancellationRequested)
                    {
                        return;
                    }

                    if (State.Status == ParserState.Error) { return; }

                    var qualifiedName = new QualifiedModuleName(toParse[index]);
                    Logger.Debug("Module '{0}' {1}", qualifiedName.ComponentName,
                        State.IsNewOrModified(qualifiedName) ? "was modified" : "was NOT modified");

                    lock (State)
                    {
                        State.SetModuleState(toParse[index], ParserState.ResolvingDeclarations);
                    }

                    ResolveDeclarations(qualifiedName.Component,
                        State.ParseTrees.Find(s => s.Key == qualifiedName).Value);
                });

                parseTasks[i].Start();
            }

            Task.WaitAll(parseTasks);

            if (State.Status < ParserState.Error)
            {
                State.SetStatusAndFireStateChanged(ParserState.ResolvedDeclarations);
                ResolveReferencesAsync();
            }
        }

        private Task[] ResolveReferencesAsync()
        {
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
                if (_resolverTokenSource.IsCancellationRequested || _central.IsCancellationRequested)
                {
                    return new Task[0];
                }

                tasks[index] = Task.Run(() =>
                {
                    lock (State)
                    {
                        State.SetModuleState(kvp.Key.Component, ParserState.ResolvingReferences);
                    }

                    ResolveReferences(finder, kvp.Key.Component, kvp.Value);
                });
            }

            return tasks;
        }

        private void AddBuiltInDeclarations()
        {
            lock (State)
            {
                foreach (var customDeclarationLoader in _customDeclarationLoaders)
                {
                    foreach (var declaration in customDeclarationLoader.Load())
                    {
                        State.AddDeclaration(declaration);
                    }
                }
            }
        }

        private readonly HashSet<ReferencePriorityMap> _projectReferences = new HashSet<ReferencePriorityMap>();

        private string GetReferenceProjectId(Reference reference, IReadOnlyList<VBProject> projects)
        {
            VBProject project = null;
            foreach (var item in projects)
            {
                try
                {
                    if (item.FileName == reference.FullPath)
                    {
                        project = item;
                    }
                }
                catch (IOException)
                {
                    // Filename throws exception if unsaved.
                }
            }
            
            if (project != null)
            {
                return QualifiedModuleName.GetProjectId(project);
            }
            return QualifiedModuleName.GetProjectId(reference);
        }

        private void SyncComReferences(IReadOnlyList<VBProject> projects)
        {
            var loadTasks = new List<Task>();

            foreach (var vbProject in projects)
            {
                var projectId = QualifiedModuleName.GetProjectId(vbProject);
                // use a 'for' loop to store the order of references as a 'priority'.
                // reference resolver needs this to know which declaration to prioritize when a global identifier exists in multiple libraries.
                for (var priority = 1; priority <= vbProject.References.Count; priority++)
                {
                    var reference = vbProject.References.Item(priority);
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
                        map = new ReferencePriorityMap(referencedProjectId) { { projectId, priority } };
                        _projectReferences.Add(map);
                    }
                    else
                    {
                        map[projectId] = priority;
                    }

                    if (!map.IsLoaded)
                    {
                        State.OnStatusMessageUpdate(ParserState.LoadingReference.ToString());

                        var tightlyScopedCapture = reference;

                        loadTasks.Add(
                        Task.Run(() =>
                        {
                            var comReflector = new ReferencedDeclarationsCollector(State);
                            var items = comReflector.GetDeclarationsForReference(tightlyScopedCapture);

                            foreach (var declaration in items)
                            {
                                State.AddDeclaration(declaration);
                            }
                        }));
                        map.IsLoaded = true;
                    }
                }
            }

            var mappedIds = new List<string>();
            foreach (var item in _projectReferences)
            {
                mappedIds.Add(item.ReferencedProjectId);
            }

            var unmapped = new List<Reference>();
            foreach (var project in projects)
            {
                foreach (Reference item in project.References)
                {
                    if (!mappedIds.Contains(GetReferenceProjectId(item, projects)))
                    {
                        unmapped.Add(item);
                    }
                }
            }

            Task.WaitAll(loadTasks.ToArray());

            foreach (var reference in unmapped)
            {
                UnloadComReference(reference, projects);
            }
        }

        private void UnloadComReference(Reference reference, IReadOnlyList<VBProject> projects)
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
                Debug.Assert(false);
                return;
            }
            map.Remove(referencedProjectId);
            if (map.Count == 0)
            {
                _projectReferences.Remove(map);
                State.RemoveBuiltInDeclarations(reference);
            }
        }

        private Task ParseAsync(VBComponent component, CancellationToken token, TokenStreamRewriter rewriter = null)
        {
            lock (State)
                lock (component)
                {
                    State.ClearStateCache(component);
                    State.SetModuleState(component, ParserState.Pending); // also clears module-exceptions
                }

            var linkedTokenSource = CancellationTokenSource.CreateLinkedTokenSource(_central.Token, token);

            var task = new Task(() => ParseAsyncInternal(component, linkedTokenSource.Token, rewriter));
            _currentTasks.TryAdd(component, Tuple.Create(task, linkedTokenSource));

            Tuple<Task, CancellationTokenSource> removedTask;
            task.ContinueWith(t => _currentTasks.TryRemove(component, out removedTask)); // default also executes on cancel
            // See http://stackoverflow.com/questions/6800705/why-is-taskscheduler-current-the-default-taskscheduler
            task.Start(TaskScheduler.Default);
            return task;
        }

        public void Cancel(VBComponent component = null)
        {
            lock (_central)
                lock (_resolverTokenSource)
                {
                    if (component == null)
                    {
                        _central.Cancel(false);

                        _central.Dispose();
                        _central = new CancellationTokenSource();
                        _resolverTokenSource = CancellationTokenSource.CreateLinkedTokenSource(_central.Token);
                    }
                    else
                    {
                        _resolverTokenSource.Cancel(false);
                        _resolverTokenSource.Dispose();

                        _resolverTokenSource = CancellationTokenSource.CreateLinkedTokenSource(_central.Token);
                        Tuple<Task, CancellationTokenSource> result;
                        if (_currentTasks.TryGetValue(component, out result))
                        {
                            result.Item2.Cancel(false);
                            result.Item2.Dispose();
                        }
                    }
                }
        }

        private void ParseAsyncInternal(VBComponent component, CancellationToken token, TokenStreamRewriter rewriter = null)
        {
            var preprocessor = _preprocessorFactory();
            var parser = new ComponentParseTask(component, preprocessor, _attributeParser, rewriter);
            parser.ParseFailure += (sender, e) =>
            {
                lock (State)
                    lock (component)
                    {
                        State.SetModuleState(component, ParserState.Error, e.Cause as SyntaxErrorException);
                    }
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
            lock (State)
                lock (component)
                {
                    State.SetModuleState(component, ParserState.Parsing);
                }
            parser.Start(token);
        }

        private readonly ConcurrentDictionary<string, Declaration> _projectDeclarations = new ConcurrentDictionary<string, Declaration>();
        private void ResolveDeclarations(VBComponent component, IParseTree tree)
        {
            if (component == null) { return; }

            var qualifiedModuleName = new QualifiedModuleName(component);

            try
            {
                var project = component.Collection.Parent;
                var projectQualifiedName = new QualifiedModuleName(project);
                Declaration projectDeclaration;
                if (!_projectDeclarations.TryGetValue(projectQualifiedName.ProjectId, out projectDeclaration))
                {
                    projectDeclaration = CreateProjectDeclaration(projectQualifiedName, project);
                    _projectDeclarations.AddOrUpdate(projectQualifiedName.ProjectId, projectDeclaration, (s, c) => projectDeclaration);
                    lock (State)
                    {
                        State.AddDeclaration(projectDeclaration);
                    }
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
                lock (State)
                {
                    State.SetModuleState(component, ParserState.ResolverError);
                }
            }
        }

        private Declaration CreateProjectDeclaration(QualifiedModuleName projectQualifiedName, VBProject project)
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

        private void ResolveReferences(DeclarationFinder finder, VBComponent component, IParseTree tree)
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
                    Logger.Debug("Binding Resolution done for component '{0}' in {1}ms (thread {2})", component.Name,
                        watch.ElapsedMilliseconds, Thread.CurrentThread.ManagedThreadId);

                    State.RebuildSelectionCache();
                    State.SetModuleState(component, ParserState.Ready);
                }
                catch (Exception exception)
                {
                    Logger.Error(exception, "Exception thrown resolving '{0}' (thread {1}).", component.Name, Thread.CurrentThread.ManagedThreadId);
                    State.SetModuleState(component, ParserState.ResolverError);
                }
            }
            
            Logger.Debug("'{0}' is {1} (thread {2})", component.Name, State.GetModuleState(component), Thread.CurrentThread.ManagedThreadId);
        }

        public void Dispose()
        {
            State.ParseRequest -= ReparseRequested;

            if (_central != null)
            {
                //_central.Cancel();
                _central.Dispose();
            }

            if (_resolverTokenSource != null)
            {
                _resolverTokenSource.Dispose();
            }
        }
    }
}
