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
        public RubberduckParser(VBE vbe, RubberduckParserState state)
        {
            _vbe = vbe;
            _state = state;

            state.ParseRequest += ReparseRequested;
        }

        private void ReparseRequested(object sender, EventArgs e)
        {
            Task.Run(() => ParseParallel());
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
                    ParseComponent(component, false);
                }

                using (var tokenSource = new CancellationTokenSource())
                {
                    Resolve(tokenSource.Token);
                }
            }
            catch (Exception exception)
            {
                Debug.Print(exception.ToString());
            }
        }

        public void ParseComponent(VBComponent component, bool resolve = true, TokenStreamRewriter rewriter = null)
        {
            var tokenSource = RenewTokenSource(component);

            var token = tokenSource.Token;
            Parse(component, token, rewriter);

            // don't fire up the resolver if any component is still being parsed
            if (resolve && _state.Status == ParserState.Parsed && !token.IsCancellationRequested)
            {
                using (var source = new CancellationTokenSource())
                {
                    Resolve(source.Token);
                }
            }
        }

        private CancellationTokenSource RenewTokenSource(VBComponent component)
        {
            if (_tokenSources.ContainsKey(component))
            {
                CancellationTokenSource existingTokenSource;
                _tokenSources.TryRemove(component, out existingTokenSource);
                if (existingTokenSource != null)
                {
                    existingTokenSource.Cancel();
                    existingTokenSource.Dispose();
                }
            }

            var tokenSource = new CancellationTokenSource();
            _tokenSources[component] = tokenSource;
            return tokenSource;
        }

        private void ParseParallel()
        {
            try
            {
                var components = _vbe.VBProjects.Cast<VBProject>()
                    .Where(project => project.Protection == vbext_ProjectProtection.vbext_pp_none)
                    .SelectMany(project => project.VBComponents.Cast<VBComponent>())
                    .ToList();

                SetComponentsState(components, ParserState.Pending);

                var parseTasks = components.Select(vbComponent => Task.Run(() => ParseComponent(vbComponent, false))).ToArray();
                Task.WhenAll(parseTasks)
                    .ContinueWith(t =>
                    {
                        using (var tokenSource = new CancellationTokenSource())
                        {
                            Resolve(tokenSource.Token);
                        }
                    });
            }
            catch (Exception exception)
            {
                Debug.Print(exception.ToString());
            }
        }

        private void SetComponentsState(IEnumerable<VBComponent> components, ParserState state)
        {
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
                    : rewriter.GetText(); // note: removes everything ignored by the parser, e.g. line numbers and comments

                ParseInternal(component, code, token);
            }
            catch (COMException exception)
            {
                State.SetModuleState(component, ParserState.Error);
                Debug.Print(exception.ToString());

            }
            catch (SyntaxErrorException exception)
            {
                Debug.Print(exception.ToString());
                State.SetModuleState(component, ParserState.Error, exception);
            }
        }

        public void Resolve(CancellationToken token)
        {
            try
            {
                var resolverTasks = _state.ParseTrees.Select(kvp => Task.Run(() =>
                {
                    token.ThrowIfCancellationRequested();
                    ResolveReferences(kvp.Key, kvp.Value, token);
                }, token)).ToArray();

                Task.WaitAll(resolverTasks);
            }
            catch (OperationCanceledException)
            {
                // let it go...
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
            var tree = ParseInternal(code, listeners, out stream);
            _state.AddTokenStream(vbComponent, stream);
            _state.AddParseTree(vbComponent, tree);

            token.ThrowIfCancellationRequested();

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

            var comments = ParseComments(qualifiedName, commentListener.Comments, commentListener.RemComments);
            _state.SetModuleComments(vbComponent, comments);

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

        private void ResolveReferences(VBComponent component, IParseTree tree, CancellationToken token)
        {
            if (_state.GetModuleState(component) != ParserState.Parsed)
            {
                return;
            }

            Debug.Print("Resolving '{0}'...", component.Name);

            if (!string.IsNullOrWhiteSpace(tree.GetText().Trim()))
            {
                var resolver = new IdentifierReferenceResolver(new QualifiedModuleName(component), _state.AllDeclarations, _state.AllComments);
                var listener = new IdentifierReferenceListener(resolver, token);
                var walker = new ParseTreeWalker();
                try
                {
                    walker.Walk(listener, tree);
                }
                catch (Exception exception)
                {
                    Debug.Print("Exception thrown resolving '{0}': {1}", component.Name, exception);
                }
            }
            _state.SetModuleState(component, ParserState.Ready);
            Debug.Print("'{0}' is {1}.", component.Name, _state.GetModuleState(component));
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
