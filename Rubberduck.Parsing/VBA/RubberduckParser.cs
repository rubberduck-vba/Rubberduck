using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System.Globalization;
using Rubberduck.Parsing.Preprocessing;
using System.Diagnostics;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Nodes;
using Rubberduck.VBEditor.Extensions;
using System.IO;

namespace Rubberduck.Parsing.VBA
{
    public class RubberduckParser : IRubberduckParser
    {
        public RubberduckParserState State
        {
            get
            {
                return _state;
            }
        }

        private CancellationTokenSource _central = new CancellationTokenSource();
        private CancellationTokenSource _resolverTokenSource; // linked to _central later
        private readonly ConcurrentDictionary<VBComponent, Tuple<Task, CancellationTokenSource>> _currentTasks =
            new ConcurrentDictionary<VBComponent, Tuple<Task, CancellationTokenSource>>();

        private readonly Dictionary<VBComponent, IParseTree> _parseTrees = new Dictionary<VBComponent, IParseTree>();
        private readonly Dictionary<QualifiedModuleName, Dictionary<Declaration, byte>> _declarations = new Dictionary<QualifiedModuleName, Dictionary<Declaration, byte>>();
        private readonly Dictionary<VBComponent, ITokenStream> _tokenStreams = new Dictionary<VBComponent, ITokenStream>();
        private readonly Dictionary<VBComponent, IList<CommentNode>> _comments = new Dictionary<VBComponent, IList<CommentNode>>();
        private readonly IDictionary<VBComponent, IDictionary<Tuple<string, DeclarationType>, Attributes>> _componentAttributes
            = new Dictionary<VBComponent, IDictionary<Tuple<string, DeclarationType>, Attributes>>();


        private readonly ReferencedDeclarationsCollector _comReflector;

        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IAttributeParser _attributeParser;

        public RubberduckParser(VBE vbe, RubberduckParserState state, IAttributeParser attributeParser)
        {
            _resolverTokenSource = CancellationTokenSource.CreateLinkedTokenSource(_central.Token);
            _vbe = vbe;
            _state = state;
            _attributeParser = attributeParser;

            _comReflector = new ReferencedDeclarationsCollector();

            state.ParseRequest += ReparseRequested;
            state.StateChanged += StateOnStateChanged;
        }

        private void StateOnStateChanged(object sender, EventArgs e)
        {
            Debug.WriteLine("RubberduckParser handles OnStateChanged ({0})", _state.Status);

            if (_state.Status == ParserState.Parsed)
            {
                Debug.WriteLine("(handling OnStateChanged) Starting resolver task");
                Resolve(_central.Token); // Tests expect this to be synchronous
            }
        }

        private void ReparseRequested(object sender, ParseRequestEventArgs e)
        {
            if (e.IsFullReparseRequest)
            {
                Cancel();
                ParseAll();
            }
            else
            {
                Cancel(e.Component);
                ParseAsync(e.Component, CancellationToken.None);
            }
        }

        public void Parse()
        {
            if (!_state.Projects.Any())
            {
                foreach (var project in _vbe.VBProjects.Cast<VBProject>())
                {
                    _state.AddProject(project);
                }
            }

            var projects = _state.Projects
                .Where(project => project.Protection == vbext_ProjectProtection.vbext_pp_none)
                .ToList();

            var components = projects.SelectMany(p => p.VBComponents.Cast<VBComponent>()).ToList();
            SyncComReferences(projects);

            foreach (var component in components)
            {
                _state.SetModuleState(component, ParserState.Pending);
            }

            // invalidation cleanup should go into ParseAsync?
            foreach (var invalidated in _componentAttributes.Keys.Except(components))
            {
                _componentAttributes.Remove(invalidated);
            }

            foreach (var vbComponent in components)
            {
                while (!_state.ClearDeclarations(vbComponent)) { }

                // expects synchronous parse :/
                ParseComponent(vbComponent);
            }
        }

        /// <summary>
        /// Starts parsing all components of all unprotected VBProjects associated with the VBE-Instance passed to the constructor of this parser instance.
        /// </summary>
        private void ParseAll()
        {
            var projects = _state.Projects
                // accessing the code of a protected VBComponent throws a COMException:
                .Where(project => project.Protection == vbext_ProjectProtection.vbext_pp_none)
                .ToList();

            var components = projects.SelectMany(p => p.VBComponents.Cast<VBComponent>()).ToList();
            var modified = components.Where(_state.IsModified).ToList();
            var unchanged = components.Where(c => !_state.IsModified(c)).ToList();

            SyncComReferences(projects);

            if (!modified.Any())
            {
                return;
            }

            foreach (var component in modified)
            {
                _state.SetModuleState(component, ParserState.Pending);
            }
            foreach (var component in unchanged)
            {
                _state.SetModuleState(component, ParserState.Parsed);
            }

            // invalidation cleanup should go into ParseAsync?
            foreach (var invalidated in _componentAttributes.Keys.Except(components))
            {
                _componentAttributes.Remove(invalidated);
            }

            foreach (var vbComponent in modified)
            {
                ParseAsync(vbComponent, CancellationToken.None);
            }
        }

        private readonly HashSet<ReferencePriorityMap> _projectReferences = new HashSet<ReferencePriorityMap>();

        private string GetReferenceProjectId(Reference reference, IReadOnlyList<VBProject> projects)
        {
            var id = projects.FirstOrDefault(project =>
            {
                try
                {
                    return project.FileName == reference.FullPath;
                }
                catch(IOException)
                {
                    // Filename throws exception if unsaved.
                    return false;
                }
            });
            if (id != null)
            {
                return QualifiedModuleName.GetProjectId(id);
            }
            return QualifiedModuleName.GetProjectId(reference);
        }

        private void SyncComReferences(IReadOnlyList<VBProject> projects)
        {
            foreach (var vbProject in projects)
            {
                var projectId = QualifiedModuleName.GetProjectId(vbProject);
                for (var priority = 1; priority <= vbProject.References.Count; priority++)
                {
                    var reference = vbProject.References.Item(priority);
                    var referencedProjectId = GetReferenceProjectId(reference, projects);
                    var map = _projectReferences.SingleOrDefault(r => r.ReferencedProjectId == referencedProjectId);
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
                        var items = _comReflector.GetDeclarationsForReference(reference).ToList();
                        foreach (var declaration in items)
                        {
                            _state.AddDeclaration(declaration);
                        }
                        map.IsLoaded = true;
                    }
                }
            }

            var mappedIds = _projectReferences.Select(map => map.ReferencedProjectId);
            var unmapped = projects.SelectMany(project => project.References.Cast<Reference>())
                .Where(reference => !mappedIds.Contains(GetReferenceProjectId(reference, projects)));
            foreach (var reference in unmapped)
            {
                UnloadComReference(reference, projects);
            }
        }

        private void UnloadComReference(Reference reference, IReadOnlyList<VBProject> projects)
        {
            var referencedProjectId = GetReferenceProjectId(reference, projects);
            var map = _projectReferences.SingleOrDefault(r => r.ReferencedProjectId == referencedProjectId);
            if (map == null || !map.IsLoaded)
            {
                // we're removing a reference we weren't tracking? ...this shouldn't happen.
                Debug.Assert(false);
                return;
            }
            map.Remove(referencedProjectId);
            if (!map.Any())
            {
                _projectReferences.Remove(map);
                _state.RemoveBuiltInDeclarations(reference);
            }
        }

        public Task ParseAsync(VBComponent component, CancellationToken token, TokenStreamRewriter rewriter = null)
        {
            _state.ClearDeclarations(component);
            _state.SetModuleState(component, ParserState.Pending); // also clears module-exceptions

            var linkedTokenSource = CancellationTokenSource.CreateLinkedTokenSource(_central.Token, token);

            //var taskFactory = new TaskFactory(new StaTaskScheduler());
            var task = new Task(() => ParseAsyncInternal(component, linkedTokenSource.Token, rewriter));
            _currentTasks.TryAdd(component, Tuple.Create(task, linkedTokenSource));

            Tuple<Task, CancellationTokenSource> removedTask;
            task.ContinueWith(t => _currentTasks.TryRemove(component, out removedTask)); // default also executes on cancel

            task.Start(/*taskFactory.Scheduler*/);
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
            var preprocessor = new VBAPreprocessor(double.Parse(_vbe.Version, CultureInfo.InvariantCulture));
            var parser = new ComponentParseTask(component, preprocessor, _attributeParser, rewriter);
            parser.ParseFailure += (sender, e) => _state.SetModuleState(component, ParserState.Error, e.Cause as SyntaxErrorException);
            parser.ParseCompleted += (sender, e) =>
            {
                // possibly lock _state
                _state.SetModuleAttributes(component, e.Attributes);
                _state.AddParseTree(component, e.ParseTree);
                _state.AddTokenStream(component, e.Tokens);
                _state.SetModuleComments(component, e.Comments);
                _state.SetModuleAnnotations(component, e.Annotations);

                // This really needs to go last
                _state.SetModuleState(component, ParserState.Parsed);
            };
            _state.SetModuleState(component, ParserState.Parsing);
            parser.Start(token);
        }

        private void ParseComponent(VBComponent component, TokenStreamRewriter rewriter = null)
        {
            ParseAsync(component, CancellationToken.None, rewriter).Wait();
        }

        private void Resolve(CancellationToken token)
        {
            var sharedTokenSource = CancellationTokenSource.CreateLinkedTokenSource(_resolverTokenSource.Token, token);
            // tests expect this to be synchronous :/
            //Task.Run(() => ResolveInternal(sharedTokenSource.Token));
            ResolveInternal(sharedTokenSource.Token);
        }

        private void ResolveInternal(CancellationToken token)
        {
            var components = _state.Projects
                .Where(project => project.Protection == vbext_ProjectProtection.vbext_pp_none)
                .SelectMany(p => p.VBComponents.Cast<VBComponent>()).ToList();
            if (!_state.HasAllParseTrees(components))
            {
                return;
            }
            _projectDeclarations.Clear();
            foreach (var kvp in _state.ParseTrees)
            {
                var qualifiedName = kvp.Key;
                if (true /*_state.IsModified(qualifiedName)*/)
                {
                    Debug.WriteLine("Module '{0}' {1}", qualifiedName.ComponentName, _state.IsModified(qualifiedName) ? "was modified" : "was NOT modified");
                    // modified module; walk parse tree and re-acquire all declarations
                    if (token.IsCancellationRequested) return;
                    ResolveDeclarations(qualifiedName.Component, kvp.Value);
                }
                else
                {
                    Debug.WriteLine(string.Format("Module '{0}' was not modified since last parse. Clearing identifier references...", kvp.Key.ComponentName));
                    // clear identifier references for non-modified modules
                    var declarations = _state.AllUserDeclarations.Where(item => item.QualifiedName.QualifiedModuleName.Equals(qualifiedName));
                    foreach (var declaration in declarations)
                    {
                        declaration.ClearReferences();
                    }
                }
            }

            // walk all parse trees (modified or not) for identifier references
            var finder = new DeclarationFinder(_state.AllDeclarations, _state.AllComments, _state.AllAnnotations);
            foreach (var kvp in _state.ParseTrees)
            {
                if (token.IsCancellationRequested) return;
                ResolveReferences(finder, kvp.Key.Component, kvp.Value);
            }
        }

        private readonly Dictionary<string, Declaration> _projectDeclarations = new Dictionary<string, Declaration>(); 
        private void ResolveDeclarations(VBComponent component, IParseTree tree)
        {
            var qualifiedModuleName = new QualifiedModuleName(component);

            var obsoleteCallStatementListener = new ObsoleteCallStatementListener();
            var obsoleteLetStatementListener = new ObsoleteLetStatementListener();
            var emptyStringLiteralListener = new EmptyStringLiteralListener();
            var argListWithOneByRefParamListener = new ArgListWithOneByRefParamListener();

            try
            {
                ParseTreeWalker.Default.Walk(new CombinedParseTreeListener(new IParseTreeListener[]{
                    obsoleteCallStatementListener,
                    obsoleteLetStatementListener,
                    emptyStringLiteralListener,
                    argListWithOneByRefParamListener,
                }), tree);
                // TODO: these are actually (almost) isnpection results.. we should handle them as such
                _state.ArgListsWithOneByRefParam = argListWithOneByRefParamListener.Contexts.Select(context => new QualifiedContext(qualifiedModuleName, context));
                _state.EmptyStringLiterals = emptyStringLiteralListener.Contexts.Select(context => new QualifiedContext(qualifiedModuleName, context));
                _state.ObsoleteLetContexts = obsoleteLetStatementListener.Contexts.Select(context => new QualifiedContext(qualifiedModuleName, context));
                _state.ObsoleteCallContexts = obsoleteCallStatementListener.Contexts.Select(context => new QualifiedContext(qualifiedModuleName, context));
                var project = component.Collection.Parent;
                var projectQualifiedName = new QualifiedModuleName(project);
                Declaration projectDeclaration;
                if (!_projectDeclarations.TryGetValue(projectQualifiedName.ProjectId, out projectDeclaration))
                {
                    projectDeclaration = CreateProjectDeclaration(projectQualifiedName, project);
                    _projectDeclarations.Add(projectQualifiedName.ProjectId, projectDeclaration);
                }
                var declarationsListener = new DeclarationSymbolsListener(qualifiedModuleName, Accessibility.Implicit, component.Type, _state.GetModuleComments(component), _state.GetModuleAnnotations(component), _state.GetModuleAttributes(component), _projectReferences, projectDeclaration);
                // TODO: should we unify the API? consider working like the other listeners instead of event-based
                declarationsListener.NewDeclaration += (sender, e) => _state.AddDeclaration(e.Declaration);
                declarationsListener.CreateModuleDeclarations();
                // rewalk parse tree for second declaration level
                ParseTreeWalker.Default.Walk(declarationsListener, tree);
            }
            catch (Exception exception)
            {
                Debug.Print("Exception thrown resolving '{0}' (thread {2}): {1}", component.Name, exception, Thread.CurrentThread.ManagedThreadId);
                _state.SetModuleState(component, ParserState.ResolverError);
            }
        }

        private Declaration CreateProjectDeclaration(QualifiedModuleName projectQualifiedName, VBProject project)
        {
            var qualifiedName = projectQualifiedName.QualifyMemberName(project.Name);
            var projectId = qualifiedName.QualifiedModuleName.ProjectId;
            var projectDeclaration = new ProjectDeclaration(qualifiedName, project.Name);
            var references = _projectReferences.Where(projectContainingReference => projectContainingReference.ContainsKey(projectId));
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

            Debug.WriteLine("Resolving '{0}'... (thread {1})", component.Name, Thread.CurrentThread.ManagedThreadId);
            var qualifiedName = new QualifiedModuleName(component);
            var resolver = new IdentifierReferenceResolver(qualifiedName, finder);
            var listener = new IdentifierReferenceListener(resolver);
            if (!string.IsNullOrWhiteSpace(tree.GetText().Trim()))
            {
                var walker = new ParseTreeWalker();
                try
                {
                    walker.Walk(listener, tree);
                    state = ParserState.Ready;
                }
                catch (Exception exception)
                {
                    Debug.Print("Exception thrown resolving '{0}' (thread {2}): {1}", component.Name, exception, Thread.CurrentThread.ManagedThreadId);
                    state = ParserState.ResolverError;
                }
            }

            _state.SetModuleState(component, state);
            Debug.Print("'{0}' is {1}. Resolver took {2}ms to complete (thread {3})", component.Name, _state.GetModuleState(component), /*_resolverTimer[component].ElapsedMilliseconds*/0, Thread.CurrentThread.ManagedThreadId);
        }

        #region Listener classes
        private class ObsoleteCallStatementListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.ExplicitCallStmtContext> _contexts = new List<VBAParser.ExplicitCallStmtContext>();
            public IEnumerable<VBAParser.ExplicitCallStmtContext> Contexts { get { return _contexts; } }

            public override void ExitExplicitCallStmt(VBAParser.ExplicitCallStmtContext context)
            {
                var procedureCall = context.eCS_ProcedureCall();
                if (procedureCall != null)
                {
                    if (procedureCall.CALL() != null)
                    {
                        _contexts.Add(context);
                        return;
                    }
                }

                var memberCall = context.eCS_MemberProcedureCall();
                if (memberCall == null) return;
                if (memberCall.CALL() == null) return;
                _contexts.Add(context);
            }
        }

        private class ObsoleteLetStatementListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.LetStmtContext> _contexts = new List<VBAParser.LetStmtContext>();
            public IEnumerable<VBAParser.LetStmtContext> Contexts { get { return _contexts; } }

            public override void ExitLetStmt(VBAParser.LetStmtContext context)
            {
                if (context.LET() != null)
                {
                    _contexts.Add(context);
                }
            }
        }

        private class EmptyStringLiteralListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.LiteralContext> _contexts = new List<VBAParser.LiteralContext>();
            public IEnumerable<VBAParser.LiteralContext> Contexts { get { return _contexts; } }

            public override void ExitLiteral(VBAParser.LiteralContext context)
            {
                var literal = context.STRINGLITERAL();
                if (literal != null && literal.GetText() == "\"\"")
                {
                    _contexts.Add(context);
                }
            }
        }

        private class ArgListWithOneByRefParamListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.ArgListContext> _contexts = new List<VBAParser.ArgListContext>();
            public IEnumerable<VBAParser.ArgListContext> Contexts { get { return _contexts; } }

            public override void ExitArgList(VBAParser.ArgListContext context)
            {
                if (context.arg() != null && context.arg().Count(a => a.BYREF() != null || (a.BYREF() == null && a.BYVAL() == null)) == 1)
                {
                    _contexts.Add(context);
                }
            }
        }

        #endregion
    }
}
