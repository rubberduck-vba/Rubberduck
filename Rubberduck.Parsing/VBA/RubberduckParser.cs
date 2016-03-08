using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Nodes;
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

        private async void ReparseRequested(object sender, EventArgs e)
        {
            Debug.WriteLine("{0} ({1}) requested a reparse", sender, sender.GetHashCode());
            await ParseParallel();
            Debug.WriteLine("Reparse requested by {0} ({1}) completed.", sender, sender.GetHashCode());
        }

        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        public RubberduckParserState State { get { return _state; } }

        private readonly ConcurrentDictionary<VBComponent, CancellationTokenSource> _tokenSources =
           new ConcurrentDictionary<VBComponent, CancellationTokenSource>();

        public void Parse()
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
                Debug.Print(exception.ToString());
            }
        }

        private void StateOnStateChanged(object sender, ParserStateEventArgs parserStateEventArgs)
        {
            if (parserStateEventArgs.State == ParserState.Parsed)
            {
                var finder = new DeclarationFinder(_state.AllDeclarations, _state.AllComments);
                using (var tokenSource = new CancellationTokenSource())
                {
                    Resolve(tokenSource.Token, finder);
                }
            }
        }

        public void ParseComponent(VBComponent component, TokenStreamRewriter rewriter = null)
        {
            Debug.Print("ParseComponent({0}) (Thread {1})", component.Name, Thread.CurrentThread.ManagedThreadId);
            var tokenSource = RenewTokenSource(component);

            var token = tokenSource.Token;
            Parse(component, token, rewriter);
        }

        private CancellationTokenSource RenewTokenSource(VBComponent component)
        {
            Debug.WriteLine("Renewing _tokenSources entry for component '{0}' (Thread {1})", component.Name, Thread.CurrentThread.ManagedThreadId);
            if (_tokenSources.ContainsKey(component))
            {
                CancellationTokenSource existingTokenSource;
                _tokenSources.TryRemove(component, out existingTokenSource);
                if (existingTokenSource != null)
                {
                    Debug.WriteLine("Cancelling source {0} in thread {1}", existingTokenSource.GetHashCode(), Thread.CurrentThread.ManagedThreadId);
                    existingTokenSource.Cancel();
                    existingTokenSource.Dispose();
                }
            }

            var tokenSource = new CancellationTokenSource();
            Debug.WriteLine("Token source for component '{0}' is now {1} (Thread {2})", component.Name, tokenSource.GetHashCode(), Thread.CurrentThread.ManagedThreadId);
            _tokenSources[component] = tokenSource;
            return tokenSource;
        }

        private async Task ParseParallel()
        {
            try
            {
                var projects = _vbe.VBProjects.Cast<VBProject>()
                    .Where(project => project.Protection == vbext_ProjectProtection.vbext_pp_none)
                    .ToList();

                if (!_state.AllDeclarations.Any(declaration => declaration.IsBuiltIn))
                {
                    SetComponentsState(projects.SelectMany(project => project.VBComponents.Cast<VBComponent>()), ParserState.LoadingReference);
                    // multiple projects can (do) have same references; avoid adding them multiple times!
                    var references = projects.SelectMany(project => project.References.Cast<Reference>())
                                             .DistinctBy(reference => reference.Guid);

                    Parallel.ForEach(references, reference =>
                    {
                        var stopwatch = Stopwatch.StartNew();
                        var declarations = _comReflector.GetDeclarationsForReference(reference);
                        foreach (var declaration in declarations)
                        {
                            _state.AddDeclaration(declaration);
                        }
                        stopwatch.Stop();
                        Debug.WriteLine("{0} declarations added in {1}ms (Thread {2})", reference.Name, stopwatch.ElapsedMilliseconds, Thread.CurrentThread.ManagedThreadId);
                    });

                    Debug.WriteLine("{0} built-in declarations added. (Thread {1})", _state.AllDeclarations.Count(d => d.IsBuiltIn), Thread.CurrentThread.ManagedThreadId);
                }

                var components = projects
                    .SelectMany(project => project.VBComponents.Cast<VBComponent>())
                    .ToList();

                foreach (var component in components)
                {
                    ParseComponent(component);
                }
                //Parallel.ForEach(components, new ParallelOptions(), component =>
                //{
                //    ParseComponent(component);
                //});
            }
            catch (Exception exception)
            {
                Debug.Print(exception.ToString());
            }
        }

        private void SetComponentsState(IEnumerable<VBComponent> components, ParserState state)
        {
            Debug.WriteLine("Setting all components to '{0}' state... (Thread {1})", state, Thread.CurrentThread.ManagedThreadId);
            foreach (var vbComponent in components)
            {
                _state.SetModuleState(vbComponent, state);
            }
        }

        private void Parse(VBComponent vbComponent, CancellationToken token, TokenStreamRewriter rewriter = null)
        {
            var component = vbComponent;

            try
            {
                token.ThrowIfCancellationRequested();

                var code = rewriter == null
                    ? string.Join(Environment.NewLine, vbComponent.CodeModule.GetSanitizedCode())
                    : rewriter.GetText();
                // note: removes everything ignored by the parser, e.g. line numbers

                ParseInternal(component, code, token);
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

        public void Resolve(CancellationToken token, DeclarationFinder finder)
        {
            try
            {
                Debug.WriteLine("Starting parallel resolver loop (thread {0})", Thread.CurrentThread.ManagedThreadId);

                var stopwatch = Stopwatch.StartNew();
                Parallel.ForEach(_state.ParseTrees, kvp =>
                {
                    token.ThrowIfCancellationRequested();
                    ResolveReferences(finder, kvp.Key, kvp.Value, token);
                });
                stopwatch.Stop();
                Debug.WriteLine("Resolver completed in {0}ms (thread {1})", stopwatch.ElapsedMilliseconds, Thread.CurrentThread.ManagedThreadId);
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
            var commentNodes = comments
                .Select(comment => new CommentNode(comment.GetComment(), Tokens.CommentMarker, new QualifiedSelection(qualifiedName, comment.GetSelection())));
            var remCommentNodes = remComments
                .Select(comment => new CommentNode(comment.GetComment(), Tokens.Rem, new QualifiedSelection(qualifiedName, comment.GetSelection())));
            var allCommentNodes = commentNodes.Union(remCommentNodes);
            return allCommentNodes;
        }

        private void ParseInternal(VBComponent vbComponent, string code, CancellationToken token)
        {
            _state.ClearDeclarations(vbComponent);
            State.SetModuleState(vbComponent, ParserState.Parsing);

            var qualifiedName = new QualifiedModuleName(vbComponent);

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

            token.ThrowIfCancellationRequested();

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

            token.ThrowIfCancellationRequested();

            // comments must be in the parser state before we start walking for declarations:
            var comments = ParseComments(qualifiedName, commentListener.Comments, commentListener.RemComments);
            _state.SetModuleComments(vbComponent, comments);

            // cannot locate declarations in one pass *the way it's currently implemented*,
            // because the context in EnterSubStmt() doesn't *yet* have child nodes when the context enters.
            // so we need to EnterAmbiguousIdentifier() and evaluate the parent instead - this *might* work.
            var declarationsListener = new DeclarationSymbolsListener(qualifiedName, Accessibility.Implicit, vbComponent.Type, _state.GetModuleComments(vbComponent), token);

            token.ThrowIfCancellationRequested();
            declarationsListener.NewDeclaration += declarationsListener_NewDeclaration;
            declarationsListener.CreateModuleDeclarations();

            token.ThrowIfCancellationRequested();
            var walker = new ParseTreeWalker();
            walker.Walk(declarationsListener, tree);
            declarationsListener.NewDeclaration -= declarationsListener_NewDeclaration;

            _state.ObsoleteCallContexts = obsoleteCallsListener.Contexts.Select(context => new QualifiedContext(qualifiedName, context));
            _state.ObsoleteLetContexts = obsoleteLetListener.Contexts.Select(context => new QualifiedContext(qualifiedName, context));
            _state.EmptyStringLiterals = emptyStringLiteralListener.Contexts.Select(context => new QualifiedContext(qualifiedName, context));
            _state.ArgListsWithOneByRefParam = argListsWithOneByRefParam.Contexts.Select(context => new QualifiedContext(qualifiedName, context));

            State.SetModuleState(vbComponent, ParserState.Parsed);
        }

        private IParseTree ParseInternal(string code, IEnumerable<IParseTreeListener> listeners, out ITokenStream outStream)
        {
            var stream = new AntlrInputStream(code);
            return ParseInternal(stream, listeners, out outStream);
        }

        private IParseTree ParseInternal(ICharStream stream, IEnumerable<IParseTreeListener> listeners, out ITokenStream outStream)
        {
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

        private void ResolveReferences(DeclarationFinder finder, VBComponent component, IParseTree tree, CancellationToken token)
        {
            var state = _state.GetModuleState(component);
            if (_state.Status == ParserState.ResolverError || state != ParserState.Parsed)
            {
                return;
            }

            _state.SetModuleState(component, ParserState.Resolving);
            _resolverTimer[component] = Stopwatch.StartNew();

            Debug.WriteLine("Resolving '{0}'... (thread {1})", component.Name, Thread.CurrentThread.ManagedThreadId);

            if (!string.IsNullOrWhiteSpace(tree.GetText().Trim()))
            {
                var qualifiedName = new QualifiedModuleName(component);
                var resolver = new IdentifierReferenceResolver(qualifiedName, finder);
                var listener = new IdentifierReferenceListener(resolver, token);
                var walker = new ParseTreeWalker();
                try
                {
                    walker.Walk(listener, tree);
                    _state.SetModuleState(component, ParserState.Ready);
                }
                catch (Exception exception)
                {
                    Debug.Print("Exception thrown resolving '{0}' in thread {2}: {1}", component.Name, exception, Thread.CurrentThread.ManagedThreadId);
                    State.SetModuleState(component, ParserState.ResolverError);
                }
            }

            _resolverTimer[component].Stop();
            Debug.Print("'{0}' is {1}. Resolver took {2}ms to complete in thread {3}", component.Name, _state.GetModuleState(component), _resolverTimer[component].ElapsedMilliseconds, Thread.CurrentThread.ManagedThreadId);
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
