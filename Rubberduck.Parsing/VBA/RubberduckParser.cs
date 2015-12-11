using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using NLog;
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
        private readonly VBE _vbe;
        private readonly Logger _logger;

        public RubberduckParser(VBE vbe, RubberduckParserState state)
        {
            _vbe = vbe;
            _state = state;
            _logger = LogManager.GetCurrentClassLogger();
        }

        private readonly RubberduckParserState _state;
        public RubberduckParserState State { get { return _state; } }

        public async Task ParseAsync(VBComponent vbComponent, CancellationToken token)
        {
            var component = vbComponent;

            var parseTask = Task.Run(() => ParseInternal(component, token), token);

            try
            {
                await parseTask;
            }
            catch (SyntaxErrorException exception)
            {
                State.SetModuleState(component, ParserState.Error, exception);
            }
            catch (OperationCanceledException)
            {
                // no need to blow up
            } 
        }

        public void Resolve(CancellationToken token)
        {
            var options = new ParallelOptions { CancellationToken = token };
            Parallel.ForEach(_state.ParseTrees, options, kvp =>
            {
                token.ThrowIfCancellationRequested();
                ResolveReferences(kvp.Key, kvp.Value, token);
            });
        }

        private IEnumerable<CommentNode> ParseComments(QualifiedModuleName qualifiedName)
        {
            var code = qualifiedName.Component.CodeModule.Code();
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

                        var result = new CommentNode(commentBuilder.ToString(), new QualifiedSelection(qualifiedName, selection));
                        commentBuilder.Clear();

                        yield return result;
                    }
                    else
                    {
                        // ignore line continuations in comment text:
                        commentBuilder.Append(line.Substring(index, commentLength).TrimStart());
                    }
                }
            }
        }

        private void ParseInternal(VBComponent vbComponent, CancellationToken token)
        {
            _state.ClearDeclarations(vbComponent);
            State.SetModuleState(vbComponent, ParserState.Parsing);

            var qualifiedName = new QualifiedModuleName(vbComponent);
            _state.SetModuleComments(vbComponent, ParseComments(qualifiedName));

            var obsoleteCallsListener = new ObsoleteCallStatementListener();
            var obsoleteLetListener = new ObsoleteLetStatementListener();

            var listeners = new IParseTreeListener[]
            {
                obsoleteCallsListener,
                obsoleteLetListener
            };

            token.ThrowIfCancellationRequested();

            ITokenStream stream;
            var code = string.Join("\r\n", vbComponent.CodeModule.Code());
            var tree = ParseInternal(code, listeners, out stream);

            token.ThrowIfCancellationRequested();
            _state.AddTokenStream(vbComponent, stream);
            _state.AddParseTree(vbComponent, tree);

            // cannot locate declarations in one pass *the way it's currently implemented*,
            // because the context in EnterSubStmt() doesn't *yet* have child nodes when the context enters.
            // so we need to EnterAmbiguousIdentifier() and evaluate the parent instead - this *might* work.
            var declarationsListener = new DeclarationSymbolsListener(qualifiedName, Accessibility.Implicit, vbComponent.Type, _state.Comments, token);

            token.ThrowIfCancellationRequested();
            declarationsListener.NewDeclaration += declarationsListener_NewDeclaration;
            declarationsListener.CreateModuleDeclarations();

            token.ThrowIfCancellationRequested();
            var walker = new ParseTreeWalker();
            walker.Walk(declarationsListener, tree);
            declarationsListener.NewDeclaration -= declarationsListener_NewDeclaration;

            _state.ObsoleteCallContexts = obsoleteCallsListener.Contexts.Select(context => new QualifiedContext(qualifiedName, context));
            _state.ObsoleteLetContexts = obsoleteLetListener.Contexts.Select(context => new QualifiedContext(qualifiedName, context));

            State.SetModuleState(vbComponent, ParserState.Parsed);
        }

        private IParseTree ParseInternal(string code, IEnumerable<IParseTreeListener> listeners, out ITokenStream outStream)
        {
            var input = new AntlrInputStream(code);
            var lexer = new VBALexer(input);
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
            var declarations = _state.AllDeclarations;

            var resolver = new IdentifierReferenceResolver(new QualifiedModuleName(component), declarations);
            var listener = new IdentifierReferenceListener(resolver, token);
            var walker = new ParseTreeWalker();
            try
            {
                walker.Walk(listener, tree);
            }
            catch(WalkerCancelledException)
            {
                // move on
            }

            _state.SetModuleState(component, ParserState.Ready);
        }

        private class ObsoleteCallStatementListener : VBABaseListener
        {
            private readonly IList<VBAParser.ExplicitCallStmtContext> _contexts = new List<VBAParser.ExplicitCallStmtContext>();
            public IEnumerable<VBAParser.ExplicitCallStmtContext> Contexts { get { return _contexts; } }

            public override void EnterExplicitCallStmt(VBAParser.ExplicitCallStmtContext context)
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

            public override void EnterLetStmt(VBAParser.LetStmtContext context)
            {
                if (context.LET() != null)
                {
                    _contexts.Add(context);
                }
            }
        }
    }
}
