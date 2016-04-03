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
                //Task.Run(() => Resolve(_central.Token));
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
            foreach (var component in components)
            {
                _state.SetModuleState(component, ParserState.LoadingReference);
            }

            if (!_state.AllDeclarations.Any(item => item.IsBuiltIn))
            {
                var references = projects.SelectMany(p => p.References.Cast<Reference>()).ToList();
                foreach (var reference in references)
                {
                    var items = _comReflector.GetDeclarationsForReference(reference);
                    foreach (var declaration in items)
                    {
                        _state.AddDeclaration(declaration);
                    }
                }
            }

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
                .Where(project => project.Protection == vbext_ProjectProtection.vbext_pp_none)
                .ToList();

            var components = projects.SelectMany(p => p.VBComponents.Cast<VBComponent>()).ToList();
            foreach (var component in components)
            {
                _state.SetModuleState(component, ParserState.LoadingReference);
            }

            if (!_state.AllDeclarations.Any(item => item.IsBuiltIn))
            {
                var references = projects.SelectMany(p => p.References.Cast<Reference>()).ToList();
                foreach (var reference in references)
                {
                    var items = _comReflector.GetDeclarationsForReference(reference);
                    foreach (var declaration in items)
                    {
                        _state.AddDeclaration(declaration);
                    }
                }
            }

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

                ParseAsync(vbComponent, CancellationToken.None);
            }
        }

        public Task ParseAsync(VBComponent component, CancellationToken token, TokenStreamRewriter rewriter = null)
        {
            // Remove invalidated "things" from _state
            // this includes: Declarations, Comments, Attributes, Exceptions, ParseTree and TokenStream
            // how that works with the Inspecion results is not quite clear
            _state.ClearDeclarations(component);
            _state.AddParseTree(component, null);
            _state.AddTokenStream(component, null);
            
            _state.SetModuleState(component, ParserState.Pending); // also clears module-exceptions
            _state.SetModuleComments(component, Enumerable.Empty<CommentNode>());
            _state.SetModuleAttributes(component, new Dictionary<Tuple<string, DeclarationType>, Attributes>());

            var linkedTokenSource = CancellationTokenSource.CreateLinkedTokenSource(_central.Token, token);

            var task = new Task(() => ParseAsyncInternal(component, linkedTokenSource.Token, rewriter));
            _currentTasks.TryAdd(component, Tuple.Create(task, linkedTokenSource));
            Tuple<Task, CancellationTokenSource> removedTask;
            task.ContinueWith(t => _currentTasks.TryRemove(component, out removedTask)); // default also executes on cancel

            task.Start();
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

                // This really needs to go last
                _state.SetModuleState(component, ParserState.Parsed);
            };
            _state.SetModuleState(component, ParserState.Parsing);
            parser.Start(token);
        }

        public void ParseComponent(VBComponent component, TokenStreamRewriter rewriter = null)
        {
            ParseAsync(component, CancellationToken.None, rewriter).Wait();
        }

        public void Resolve(CancellationToken token)
        {
            var sharedTokenSource = CancellationTokenSource.CreateLinkedTokenSource(_resolverTokenSource.Token, token);
            // tests expect this to be synchronous :/
            //Task.Run(() => ResolveInternal(sharedTokenSource.Token));
            ResolveInternal(sharedTokenSource.Token);
        }

        private void ResolveInternal(CancellationToken token)
        {
            foreach (var kvp in _state.ParseTrees)
            {
                if (token.IsCancellationRequested) return;
                ResolveDeclarations(kvp.Key, kvp.Value);
            }
            var finder = new DeclarationFinder(_state.AllDeclarations, _state.AllComments);
            foreach (var kvp in _state.ParseTrees)
            {
                if (token.IsCancellationRequested) return;
                ResolveReferences(finder, kvp.Key, kvp.Value);
            }
        }

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
                // FIXME these are actually (almost) isnpection results.. we should handle them as such
                _state.ArgListsWithOneByRefParam = argListWithOneByRefParamListener.Contexts.Select(context => new QualifiedContext(qualifiedModuleName, context));
                _state.EmptyStringLiterals = emptyStringLiteralListener.Contexts.Select(context => new QualifiedContext(qualifiedModuleName, context));
                _state.ObsoleteLetContexts = obsoleteLetStatementListener.Contexts.Select(context => new QualifiedContext(qualifiedModuleName, context));
                _state.ObsoleteCallContexts = obsoleteCallStatementListener.Contexts.Select(context => new QualifiedContext(qualifiedModuleName, context));

                // cannot locate declarations in one pass *the way it's currently implemented*,
                // because the context in EnterSubStmt() doesn't *yet* have child nodes when the context enters.
                // so we need to EnterAmbiguousIdentifier() and evaluate the parent instead - this *might* work.
                var declarationsListener = new DeclarationSymbolsListener(qualifiedModuleName, Accessibility.Implicit, component.Type, _state.GetModuleComments(component), _state.getModuleAttributes(component));
                // TODO: should we unify the API? consider working like the other listeners instead of event-based
                declarationsListener.NewDeclaration += (sender, e) => _state.AddDeclaration(e.Declaration);
                declarationsListener.CreateModuleDeclarations();
                // rewalk parse tree for second declaration level
                ParseTreeWalker.Default.Walk(declarationsListener, tree);
            } catch (Exception exception)
            {
                Debug.Print("Exception thrown resolving '{0}' (thread {2}): {1}", component.Name, exception, Thread.CurrentThread.ManagedThreadId);
                _state.SetModuleState(component, ParserState.ResolverError);
            }

        }
        
        private void ResolveReferences(DeclarationFinder finder, VBComponent component, IParseTree tree)
        {
            var state = _state.GetModuleState(component);
            if (_state.Status == ParserState.ResolverError || state != ParserState.Parsed)
            {
                return;
            }
            _state.SetModuleState(component, ParserState.Resolving);
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
