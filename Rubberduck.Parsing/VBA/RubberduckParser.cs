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
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBA;

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

        public void ParseComponent(VBComponent component, bool resolve = true, TokenStreamRewriter rewriter = null)
        {
            var tokenSource = RenewTokenSource(component);

            var token = tokenSource.Token;
            Parse(component, token, rewriter);

            if (resolve && !token.IsCancellationRequested)
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
                    .SelectMany(project => project.VBComponents.Cast<VBComponent>())
                    .ToList();

                foreach (var vbComponent in components)
                {
                    _state.SetModuleState(vbComponent, ParserState.Pending);
                }

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
                State.SetModuleState(component, ParserState.Error, exception);
            }
            catch (OperationCanceledException)
            {
                State.SetModuleState(component, ParserState.Error);
            }
            catch (Exception exception)
            {
                // break here, inspect and debug.
                throw;
            }

            return;
        }

        public void Resolve(CancellationToken token)
        {
            try
            {
                var options = new ParallelOptions { CancellationToken = token };
                Parallel.ForEach(_state.ParseTrees, options, kvp =>
                {
                    token.ThrowIfCancellationRequested();
                    ResolveReferences(kvp.Key, kvp.Value, token);
                });
            }
            catch (OperationCanceledException)
            {
                // let it go...
            }
        }

        private IEnumerable<CommentNode> ParseComments(QualifiedModuleName qualifiedName)
        {
            var result = new List<CommentNode>();

            var code = qualifiedName.Component.CodeModule.GetSanitizedCode();
            var commentBuilder = new StringBuilder();
            var continuing = false;

            var startLine = 0;
            var startColumn = 0;

            for (var i = 0; i < code.Length; i++)
            {
                var line = code[i];
                var index = 0;

                if (continuing || line.HasComment(out index))
                {
                    startLine = continuing ? startLine : i;
                    startColumn = continuing ? startColumn : index;

                    var commentLength = line.Length - index;

                    continuing = line.EndsWith("_");
                    if (!continuing)
                    {
                        commentBuilder.Append(line.Substring(index, commentLength).TrimStart());
                        var selection = new Selection(startLine + 1, startColumn + 1, i + 1, line.Length + 1);

                        var comment = new CommentNode(commentBuilder.ToString(), new QualifiedSelection(qualifiedName, selection));
                        commentBuilder.Clear();
                        result.Add(comment);
                    }
                    else
                    {
                        // ignore line continuations in comment text:
                        commentBuilder.Append(line.Substring(index, commentLength).TrimStart());
                    }
                }
            }

            return result;
        }

        private void ParseInternal(VBComponent vbComponent, string code, CancellationToken token)
        {
            _state.ClearDeclarations(vbComponent);
            State.SetModuleState(vbComponent, ParserState.Parsing);

            var qualifiedName = new QualifiedModuleName(vbComponent);
            var comments = ParseComments(qualifiedName);
            _state.SetModuleComments(vbComponent, comments);

            var obsoleteCallsListener = new ObsoleteCallStatementListener();
            var obsoleteLetListener = new ObsoleteLetStatementListener();
            var emptyStringLiteralListener = new EmptyStringLiteralListener();
            var argListsWithOneByRefParam = new ArgListWithOneByRefParamListener();

            var listeners = new IParseTreeListener[]
            {
                obsoleteCallsListener,
                obsoleteLetListener,
                emptyStringLiteralListener,
                argListsWithOneByRefParam,
            };

            token.ThrowIfCancellationRequested();

            IParseTree tree;

            try
            {
                ITokenStream stream;
                tree = ParseInternal(code, listeners, out stream);
                _state.AddTokenStream(vbComponent, stream);
                _state.AddParseTree(vbComponent, tree);
            }
            catch (SyntaxErrorException exception)
            {
                State.SetModuleState(vbComponent, ParserState.Error, exception);
                throw;
            }

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

            _state.SetModuleState(component, ParserState.Resolving);

            var resolver = new IdentifierReferenceResolver(new QualifiedModuleName(component), _state.AllDeclarations, _state.AllComments);
            var listener = new IdentifierReferenceListener(resolver, token);
            var walker = new ParseTreeWalker();
            try
            {
                walker.Walk(listener, tree);
            }
            catch (WalkerCancelledException)
            {
                // move on
            }

            _state.SetModuleState(component, ParserState.Ready);
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
    }
}
