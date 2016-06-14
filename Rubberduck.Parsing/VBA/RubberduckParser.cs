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
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
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


        private readonly ReferencedDeclarationsCollector _comReflector;

        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IAttributeParser _attributeParser;
        private readonly Func<IVBAPreprocessor> _preprocessorFactory;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public RubberduckParser(
            VBE vbe,
            RubberduckParserState state,
            IAttributeParser attributeParser,
            Func<IVBAPreprocessor> preprocessorFactory)
        {
            _resolverTokenSource = CancellationTokenSource.CreateLinkedTokenSource(_central.Token);
            _vbe = vbe;
            _state = state;
            _attributeParser = attributeParser;
            _preprocessorFactory = preprocessorFactory;

            _comReflector = new ReferencedDeclarationsCollector(_state);

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
                    ParseAsync(e.Component, CancellationToken.None).Wait();

                    if (_state.Status != ParserState.Error)
                    {
                        Logger.Trace("Starting resolver task");
                        Resolve(_central.Token);
                    }
                });
            }
        }

        /// <summary>
        /// For the use of tests only
        /// </summary>
        public void Parse()
        {
            if (_state.Projects.Count == 0)
            {
                foreach (var project in _vbe.VBProjects.UnprotectedProjects())
                {
                    _state.AddProject(project);
                }
            }

            var components = new List<VBComponent>();
            foreach (var project in _state.Projects)
            {
                foreach (VBComponent component in project.VBComponents)
                {
                    components.Add(component);
                }
            }

            SyncComReferences(_state.Projects);

            foreach (var component in components)
            {
                _state.SetModuleState(component, ParserState.Pending);
            }

            // invalidation cleanup should go into ParseAsync?
            foreach (var key in _componentAttributes.Keys)
            {
                if (!components.Contains(key))
                {
                    _componentAttributes.Remove(key);
                }
            }

            var parseTasks = new Task[components.Count];
            for (var i = 0; i < components.Count; i++)
            {
                parseTasks[i] = ParseAsync(components[i], CancellationToken.None);
            }

            Task.WaitAll(parseTasks);

            if (_state.Status != ParserState.Error)
            {
                Logger.Trace("Starting resolver task");
                Resolve(_central.Token); // Tests expect this to be synchronous
            }
        }

        /// <summary>
        /// Starts parsing all components of all unprotected VBProjects associated with the VBE-Instance passed to the constructor of this parser instance.
        /// </summary>
        private void ParseAll()
        {
            if (_state.Projects.Count == 0)
            {
                foreach (var project in _vbe.VBProjects.UnprotectedProjects())
                {
                    _state.AddProject(project);
                }
            }

            var components = new List<VBComponent>();
            foreach (var project in _state.Projects)
            {
                foreach (VBComponent component in project.VBComponents)
                {
                    components.Add(component);
                }
            }

            var toParse = new List<VBComponent>();
            var unchanged = new List<VBComponent>();

            foreach (var component in components)
            {
                if (_state.IsNewOrModified(component))
                {
                    toParse.Add(component);
                }
                else
                {
                    unchanged.Add(component);
                }
            }

            SyncComReferences(_state.Projects);
            AddBuiltInDeclarations(_state.Projects);

            if (toParse.Count == 0)
            {
                State.SetStatusAndFireStateChanged(_state.Status);
                return;
            }


            lock (_state)  // note, method is invoked from UI thread... really need the lock here?
            {
                foreach (var component in toParse)
                {
                    _state.SetModuleState(component, ParserState.Pending);
                }
                foreach (var component in unchanged)
                {
                    // note: seting to 'Parsed' would include them in the resolver walk. 'Ready' excludes them.
                    _state.SetModuleState(component, ParserState.Ready);
                }

                //Debug.Assert(unchanged.All(component => _state.GetModuleState(component) == ParserState.Ready));
                //Debug.Assert(toParse.All(component => _state.GetModuleState(component) == ParserState.Pending));
            }

            // invalidation cleanup should go into ParseAsync?
            foreach (var key in _componentAttributes.Keys)
            {
                if (!components.Contains(key))
                {
                    _componentAttributes.Remove(key);
                }
            }

            var parseTasks = new Task[components.Count];
            for (var i = 0; i < components.Count; i++)
            {
                parseTasks[i] = ParseAsync(components[i], CancellationToken.None);
            }

            Task.WaitAll(parseTasks);

            if (_state.Status != ParserState.Error)
            {
                Logger.Trace("Starting resolver task");
                Resolve(_central.Token);
            }
        }

        private void AddBuiltInDeclarations(IReadOnlyList<VBProject> projects)
        {
            var finder = new DeclarationFinder(_state.AllDeclarations, new CommentNode[] { }, new IAnnotation[] { });

            foreach (var item in finder.MatchName(Tokens.Err))
            {
                if (item.IsBuiltIn && item.DeclarationType == DeclarationType.Variable &&
                    item.Accessibility == Accessibility.Global)
                {
                    return;
                }
            }

            var vba = finder.FindProject("VBA");
            if (vba == null)
            {
                // if VBA project is null, we haven't loaded any COM references;
                // we're in a unit test and mock project didn't setup any references.
                return;
            }
            var informationModule = finder.FindStdModule("Information", vba, true);
            Debug.Assert(informationModule != null, "We expect the information module to exist in the VBA project.");
            var customDeclarations = CustomDeclarations.Load(vba, informationModule);
            lock (_state)
            {
                foreach (var customDeclaration in customDeclarations)
                {
                    _state.AddDeclaration(customDeclaration);
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
                        _state.OnStatusMessageUpdate(ParserState.LoadingReference.ToString());
                        var items = _comReflector.GetDeclarationsForReference(reference);
                        foreach (var declaration in items)
                        {
                            _state.AddDeclaration(declaration);
                        }
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
                _state.RemoveBuiltInDeclarations(reference);
            }
        }

        public Task ParseAsync(VBComponent component, CancellationToken token, TokenStreamRewriter rewriter = null)
        {
            lock (_state)
                lock (component)
                {
                    _state.ClearStateCache(component);
                    _state.SetModuleState(component, ParserState.Pending); // also clears module-exceptions
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
                lock (_state)
                    lock (component)
                    {
                        _state.SetModuleState(component, ParserState.Error, e.Cause as SyntaxErrorException);
                    }
            };
            parser.ParseCompleted += (sender, e) =>
            {
                lock (_state)
                    lock (component)
                    {
                        _state.SetModuleAttributes(component, e.Attributes);
                        _state.AddParseTree(component, e.ParseTree);
                        _state.AddTokenStream(component, e.Tokens);
                        _state.SetModuleComments(component, e.Comments);
                        _state.SetModuleAnnotations(component, e.Annotations);

                        // This really needs to go last
                        _state.SetModuleState(component, ParserState.Parsed);
                    }
            };
            lock (_state)
                lock (component)
                {
                    _state.SetModuleState(component, ParserState.Parsing);
                }
            parser.Start(token);
        }

        private void Resolve(CancellationToken token)
        {
            State.SetStatusAndFireStateChanged(ParserState.Resolving);
            var sharedTokenSource = CancellationTokenSource.CreateLinkedTokenSource(_resolverTokenSource.Token, token);
            // tests expect this to be synchronous :/
            //Task.Run(() => ResolveInternal(sharedTokenSource.Token));
            ResolveInternal(sharedTokenSource.Token);
        }

        private void ResolveInternal(CancellationToken token)
        {
            var components = new List<VBComponent>();
            foreach (var project in _state.Projects)
            {
                if (project.Protection == vbext_ProjectProtection.vbext_pp_locked)
                {
                    continue;
                }

                foreach (VBComponent component in project.VBComponents)
                {
                    components.Add(component);
                }
            }

            if (!_state.HasAllParseTrees(components))
            {
                return;
            }
            _projectDeclarations.Clear();
            _state.ClearBuiltInReferences();
            foreach (var kvp in _state.ParseTrees)
            {
                var qualifiedName = kvp.Key;
                Logger.Debug("Module '{0}' {1}", qualifiedName.ComponentName, _state.IsNewOrModified(qualifiedName) ? "was modified" : "was NOT modified");
                // modified module; walk parse tree and re-acquire all declarations
                if (token.IsCancellationRequested) return;
                ResolveDeclarations(qualifiedName.Component, kvp.Value);
            }

            _state.SetStatusAndFireStateChanged(ParserState.ResolvedDeclarations);

            // walk all parse trees (modified or not) for identifier references
            var finder = new DeclarationFinder(_state.AllDeclarations, _state.AllComments, _state.AllAnnotations);
            var passes = new List<ICompilationPass>
            {
                // This pass has to come first because the type binding resolution depends on it.
                new ProjectReferencePass(finder),
                new TypeHierarchyPass(finder, new VBAExpressionParser()),
                new TypeAnnotationPass(finder, new VBAExpressionParser())
            };
            passes.ForEach(p => p.Execute());
            foreach (var kvp in _state.ParseTrees)
            {
                if (token.IsCancellationRequested) return;
                ResolveReferences(finder, kvp.Key.Component, kvp.Value);
            }
        }

        private readonly Dictionary<string, Declaration> _projectDeclarations = new Dictionary<string, Declaration>();
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
                    _projectDeclarations.Add(projectQualifiedName.ProjectId, projectDeclaration);
                    lock (_state)
                    {
                        _state.AddDeclaration(projectDeclaration);
                    }
                }
                Logger.Debug("Creating declarations for module {0}.", qualifiedModuleName.Name);
                var declarationsListener = new DeclarationSymbolsListener(qualifiedModuleName, component.Type, _state.GetModuleComments(component), _state.GetModuleAnnotations(component), _state.GetModuleAttributes(component), _projectReferences, projectDeclaration);
                ParseTreeWalker.Default.Walk(declarationsListener, tree);
                foreach (var createdDeclaration in declarationsListener.CreatedDeclarations)
                {
                    _state.AddDeclaration(createdDeclaration);
                }
            }
            catch (Exception exception)
            {
                Logger.Error(exception, "Exception thrown acquiring declarations for '{0}' (thread {1}).", component.Name, Thread.CurrentThread.ManagedThreadId);
                lock (_state)
                {
                    _state.SetModuleState(component, ParserState.ResolverError);
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
            var state = _state.GetModuleState(component);
            if (_state.Status == ParserState.ResolverError || (state != ParserState.Parsed))
            {
                return;
            }
            var qualifiedName = new QualifiedModuleName(component);
            Logger.Debug("Resolving identifier references in '{0}'... (thread {1})", qualifiedName.Name, Thread.CurrentThread.ManagedThreadId);
            var resolver = new IdentifierReferenceResolver(qualifiedName, finder);
            var listener = new IdentifierReferenceListener(resolver);
            if (!string.IsNullOrWhiteSpace(tree.GetText().Trim()))
            {
                var walker = new ParseTreeWalker();
                try
                {
                    Stopwatch watch = Stopwatch.StartNew();
                    walker.Walk(listener, tree);
                    watch.Stop();
                    Logger.Debug("Binding Resolution done for component '{0}' in {1}ms (thread {2})", component.Name, watch.ElapsedMilliseconds, Thread.CurrentThread.ManagedThreadId);
                    _state.RebuildSelectionCache();
                    state = ParserState.Ready;
                }
                catch (Exception exception)
                {
                    Logger.Error(exception, "Exception thrown resolving '{0}' (thread {1}).", component.Name, Thread.CurrentThread.ManagedThreadId);
                    state = ParserState.ResolverError;
                }
            }

            _state.SetModuleState(component, state);
            Logger.Debug("'{0}' is {1} (thread {2})", component.Name, _state.GetModuleState(component), Thread.CurrentThread.ManagedThreadId);
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
