using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Parsing.Preprocessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Parsing.VBA
{
    public class RubberduckParser : IRubberduckParser
    {
        private readonly ReferencedDeclarationsCollector _comReflector;

        public RubberduckParser(VBE vbe, RubberduckParserState state)
        {
            _vbe = vbe;
            _state = state;

            _comReflector = new ReferencedDeclarationsCollector();

            state.ParseRequest += ReparseRequested;
            state.StateChanged += StateOnStateChanged;
        }

        private void ReparseRequested(object sender, EventArgs e)
        {
            Task.Run(() => Parse());
        }

        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        public RubberduckParserState State { get { return _state; } }

        /// <summary>
        /// This method is not part of the interface and should only be used for testing.
        /// Request a reparse using RubberduckParserState.OnParseRequested instead.
        /// </summary>
        public void ParseSynchronous()
        {
            try
            {
                var components = _vbe.VBProjects.Cast<VBProject>()
                    .SelectMany(project => project.VBComponents.Cast<VBComponent>())
                    .ToList();

                SetComponentsState(components, ParserState.Pending);

                foreach (var component in components)
                {
                    ParseComponent(component);
                }
            }
            catch (Exception exception)
            {
                Debug.WriteLine(exception);
            }
        }

        private void StateOnStateChanged(object sender, EventArgs e)
        {
            Debug.WriteLine("RubberduckParser handles OnStateChanged ({0})", _state.Status);
            if (_state.Status == ParserState.Parsed)
            {
                Debug.WriteLine("(handling OnStateChanged) Starting resolver task");
                var finder = new DeclarationFinder(_state.AllDeclarations, _state.AllComments);
                Resolve(finder);
            }
        }

        private void Parse()
        {
            var projects = _vbe.VBProjects.Cast<VBProject>()
                .Where(project => project.Protection == vbext_ProjectProtection.vbext_pp_none)
                .ToList();

            if (!_state.AllDeclarations.Any(declaration => declaration.IsBuiltIn))
            {
                AddDeclarationsFromProjectReferences(projects);
            }

            var components = projects
                .SelectMany(project => project.VBComponents.Cast<VBComponent>())
                .ToList();

            foreach (var vbComponent in components)
            {
                while (!_state.ClearDeclarations(vbComponent))
                {
                    // till hell freezes over?
                }
            }

            SetComponentsState(components, ParserState.Pending);
            foreach (var component in components)
            {
                ParseComponent(component);
            }
        }

        private void AddDeclarationsFromProjectReferences(IReadOnlyList<VBProject> projects)
        {
            SetComponentsState(projects.SelectMany(project => project.VBComponents.Cast<VBComponent>()), ParserState.LoadingReference);

            var references = projects.SelectMany(project => project.References.Cast<Reference>()).DistinctBy(reference => reference.Guid);
            foreach (var reference in references)
            {
                AddDeclarationsFromReference(reference);
            }
        }

        private void AddDeclarationsFromReference(Reference reference)
        {
            var declarations = _comReflector.GetDeclarationsForReference(reference);
            foreach (var declaration in declarations)
            {
                _state.AddDeclaration(declaration);
            }
        }

        private void SetComponentsState(IEnumerable<VBComponent> components, ParserState state)
        {
            foreach (var vbComponent in components)
            {
                _state.SetModuleState(vbComponent, state);
            }
        }

        public void ParseComponent(VBComponent vbComponent, TokenStreamRewriter rewriter = null)
        {
            var component = vbComponent;
            State.SetModuleState(vbComponent, ParserState.Parsing);

            try
            {
                var qualifiedName = new QualifiedModuleName(vbComponent);
                var code = rewriter == null ? string.Join(Environment.NewLine, vbComponent.CodeModule.GetSanitizedCode()) : rewriter.GetText();

                var preprocessor = new VBAPreprocessor(double.Parse(_vbe.Version, CultureInfo.InvariantCulture));
                string preprocessedModuleBody;
                try
                {
                    preprocessedModuleBody = preprocessor.Execute(code);
                }
                catch (VBAPreprocessorException)
                {
                    // Fall back to not doing any preprocessing at all.
                    preprocessedModuleBody = code;
                }

                var obsoleteCallsListener = new ObsoleteCallStatementListener();
                var obsoleteLetListener = new ObsoleteLetStatementListener();
                var emptyStringLiteralListener = new EmptyStringLiteralListener();
                var argListsWithOneByRefParam = new ArgListWithOneByRefParamListener();
                var commentListener = new CommentListener();

                var listeners = new IParseTreeListener[]
                {
                    obsoleteCallsListener,
                    obsoleteLetListener,
                    emptyStringLiteralListener,
                    argListsWithOneByRefParam,
                    commentListener
                };

                var tree = GetParseTree(vbComponent, listeners, preprocessedModuleBody, qualifiedName);
                WalkParseTree(vbComponent, listeners, qualifiedName, tree);

                State.SetModuleState(vbComponent, ParserState.Parsed);
            }
            catch (COMException exception)
            {
                State.SetModuleState(component, ParserState.Error);
                Debug.WriteLine("Exception thrown in thread {0}:\n{1}", Thread.CurrentThread.ManagedThreadId, exception);
            }
            catch (SyntaxErrorException exception)
            {
                Debug.WriteLine("Exception thrown in thread {0}:\n{1}", Thread.CurrentThread.ManagedThreadId, exception);
                State.SetModuleState(component, ParserState.Error, exception);
            }
            catch (OperationCanceledException exception)
            {
                Debug.WriteLine("Exception thrown in thread {0}:\n{1}", Thread.CurrentThread.ManagedThreadId, exception);
            }
        }

        private IParseTree GetParseTree(VBComponent vbComponent, IParseTreeListener[] listeners, string code, QualifiedModuleName qualifiedName)
        {
            var commentListener = listeners.OfType<CommentListener>().Single();
            ITokenStream stream;

            var stopwatch = Stopwatch.StartNew();
            var tree = ParseInternal(code, listeners, out stream);
            stopwatch.Stop();
            if (tree != null)
            {
                Debug.Print("IParseTree for component '{0}' acquired in {1}ms (thread {2})", vbComponent.Name, stopwatch.ElapsedMilliseconds, Thread.CurrentThread.ManagedThreadId);
            }

            _state.AddTokenStream(vbComponent, stream);
            _state.AddParseTree(vbComponent, tree);

            var comments = ParseComments(qualifiedName, commentListener.Comments, commentListener.RemComments);
            _state.SetModuleComments(vbComponent, comments);

            return tree;
        }
        private void WalkParseTree(VBComponent vbComponent, IReadOnlyList<IParseTreeListener> listeners, QualifiedModuleName qualifiedName, IParseTree tree)
        {
            var obsoleteCallsListener = listeners.OfType<ObsoleteCallStatementListener>().Single();
            var obsoleteLetListener = listeners.OfType<ObsoleteLetStatementListener>().Single();
            var emptyStringLiteralListener = listeners.OfType<EmptyStringLiteralListener>().Single();
            var argListsWithOneByRefParamListener = listeners.OfType<ArgListWithOneByRefParamListener>().Single();

            // cannot locate declarations in one pass *the way it's currently implemented*,
            // because the context in EnterSubStmt() doesn't *yet* have child nodes when the context enters.
            // so we need to EnterAmbiguousIdentifier() and evaluate the parent instead - this *might* work.
            var declarationsListener = new DeclarationSymbolsListener(qualifiedName, Accessibility.Implicit, vbComponent.Type, _state.GetModuleComments(vbComponent));

            declarationsListener.NewDeclaration += declarationsListener_NewDeclaration;
            declarationsListener.CreateModuleDeclarations();

            var walker = new ParseTreeWalker();
            walker.Walk(declarationsListener, tree);
            declarationsListener.NewDeclaration -= declarationsListener_NewDeclaration;

            _state.ObsoleteCallContexts = obsoleteCallsListener.Contexts.Select(context => new QualifiedContext(qualifiedName, context));
            _state.ObsoleteLetContexts = obsoleteLetListener.Contexts.Select(context => new QualifiedContext(qualifiedName, context));
            _state.EmptyStringLiterals = emptyStringLiteralListener.Contexts.Select(context => new QualifiedContext(qualifiedName, context));
            _state.ArgListsWithOneByRefParam = argListsWithOneByRefParamListener.Contexts.Select(context => new QualifiedContext(qualifiedName, context));
        }

        private void Resolve(DeclarationFinder finder)
        {
            try
            {
                foreach (var kvp in _state.ParseTrees)
                {
                    ResolveReferences(finder, kvp.Key, kvp.Value);
                }
            }
            catch (OperationCanceledException)
            {
                // let it go...
            }
            catch (AggregateException exceptions)
            {
                Debug.WriteLine(exceptions);
            }
        }

        private IEnumerable<CommentNode> ParseComments(QualifiedModuleName qualifiedName, IEnumerable<VBAParser.CommentContext> comments, IEnumerable<VBAParser.RemCommentContext> remComments)
        {
            var commentNodes = comments.Select(comment => new CommentNode(comment.GetComment(), Tokens.CommentMarker, new QualifiedSelection(qualifiedName, comment.GetSelection())));
            var remCommentNodes = remComments.Select(comment => new CommentNode(comment.GetComment(), Tokens.Rem, new QualifiedSelection(qualifiedName, comment.GetSelection())));
            var allCommentNodes = commentNodes.Union(remCommentNodes);
            return allCommentNodes;
        }

        private static IParseTree ParseInternal(string code, IEnumerable<IParseTreeListener> listeners, out ITokenStream outStream)
        {
            var stream = new AntlrInputStream(code);
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAParser(tokens);

            parser.AddErrorListener(new ExceptionErrorListener());
            foreach (var listener in listeners)
            {
                parser.AddParseListener(listener);
            }

            outStream = tokens;
            return parser.startRule();
        }

        private void declarationsListener_NewDeclaration(object sender, DeclarationEventArgs e)
        {
            _state.AddDeclaration(e.Declaration);
        }

        // todo: remove once performance is acceptable
        private readonly IDictionary<VBComponent, Stopwatch> _resolverTimer = new ConcurrentDictionary<VBComponent, Stopwatch>();

        private void ResolveReferences(DeclarationFinder finder, VBComponent component, IParseTree tree)
        {
            var state = _state.GetModuleState(component);
            if (_state.Status == ParserState.ResolverError || state != ParserState.Parsed)
            {
                return;
            }

            _state.SetModuleState(component, ParserState.Resolving);
            _resolverTimer[component] = Stopwatch.StartNew();

            Debug.WriteLine("Resolving '{0}'... (thread {1})", component.Name, Thread.CurrentThread.ManagedThreadId);

            state = WalkParseTree(component, tree, finder);
            _state.SetModuleState(component, state);

            _resolverTimer[component].Stop();
            Debug.Print("'{0}' is {1}. Resolver took {2}ms to complete (thread {3})", component.Name, _state.GetModuleState(component), _resolverTimer[component].ElapsedMilliseconds, Thread.CurrentThread.ManagedThreadId);
        }

        private ParserState WalkParseTree(VBComponent component, IParseTree tree, DeclarationFinder finder)
        {
            if (!string.IsNullOrWhiteSpace(tree.GetText().Trim()))
            {
                var qualifiedName = new QualifiedModuleName(component);
                var resolver = new IdentifierReferenceResolver(qualifiedName, finder);
                var listener = new IdentifierReferenceListener(resolver);
                var walker = new ParseTreeWalker();
                try
                {
                    walker.Walk(listener, tree);
                }
                catch (Exception exception)
                {
                    Debug.Print("Exception thrown resolving '{0}' (thread {2}): {1}", component.Name, exception, Thread.CurrentThread.ManagedThreadId);
                    return ParserState.ResolverError;
                }
            }
            return ParserState.Ready;
        }

        private class ObsoleteCallStatementListener : VBABaseListener
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

        private class ObsoleteLetStatementListener : VBABaseListener
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

        private class EmptyStringLiteralListener : VBABaseListener
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

        private class ArgListWithOneByRefParamListener : VBABaseListener
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

        private class CommentListener : VBABaseListener
        {
            private readonly IList<VBAParser.RemCommentContext> _remComments = new List<VBAParser.RemCommentContext>();
            public IEnumerable<VBAParser.RemCommentContext> RemComments { get { return _remComments; } }

            private readonly IList<VBAParser.CommentContext> _comments = new List<VBAParser.CommentContext>();
            public IEnumerable<VBAParser.CommentContext> Comments { get { return _comments; } }

            public override void ExitRemComment([NotNull] VBAParser.RemCommentContext context)
            {
                _remComments.Add(context);
            }

            public override void ExitComment([NotNull] VBAParser.CommentContext context)
            {
                _comments.Add(context);
            }
        }
    }
}