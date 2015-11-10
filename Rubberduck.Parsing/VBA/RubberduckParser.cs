using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
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

namespace Rubberduck.Parsing.VBA
{
    public class RubberduckParser : IRubberduckParser
    {
        private readonly VBE _vbe;
        private readonly ParseCoordinator _coordinator;
        private readonly Logger _logger;

        public RubberduckParser(VBE vbe)
        {
            _vbe = vbe;
            _logger = LogManager.GetCurrentClassLogger();
            _coordinator = new ParseCoordinator(state => State.Status = state);
        }

        private readonly RubberduckParserState _state = new RubberduckParserState();
        public RubberduckParserState State { get { return _state; } }

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

        private IParseTree Parse(string code, IEnumerable<IParseTreeListener> listeners, out ITokenStream outStream)
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

            try
            {
                _state.Exception = null;
                return parser.startRule();
            }
            catch (SyntaxErrorException e)
            {
                _state.Exception = e;
            }

            _state.Status = RubberduckParserState.State.Error;
            return null;
        }

        private Tuple<IParseTree, ITokenStream, Action> Parse(VBComponent component, IEnumerable<IParseTreeListener> listeners, CancellationToken token)
        {
            ITokenStream stream;
            var code = component.CodeModule.Lines();
            var tree = Parse(code, listeners, out stream);
            return Tuple.Create(tree, stream, new Action(() => { ResolveReferences(tree, token); }));
        }

        public void Parse(VBE vbe)
        {
            var components = vbe.VBProjects
                .Cast<VBProject>()
                .SelectMany(project => project.VBComponents.Cast<VBComponent>());

            _state.AddBuiltInDeclarations(_vbe.HostApplication());
            foreach (var vbComponent in components)
            {
                Parse(vbComponent);
            }
        }

        public void Parse(VBComponent vbComponent)
        {
            _coordinator.Start(vbComponent, () => ParseInternal(vbComponent));
        }

        private Action ParseInternal(VBComponent vbComponent)
        {
            var token = CancelPreviousTask(vbComponent);
            _state.ClearDeclarations(vbComponent);

            var qualifiedName = new QualifiedModuleName(vbComponent);
            _state.Comments = ParseComments(qualifiedName);

            var obsoleteCallsListener = new ObsoleteCallStatementListener();
            var obsoleteLetListener = new ObsoleteLetStatementListener();

            var listeners = new IParseTreeListener[]
            {
                obsoleteCallsListener,
                obsoleteLetListener
            };

            _state.Status = RubberduckParserState.State.Parsing;
            var result = Parse(vbComponent, listeners, token);

            if (result.Item1 == null)
            {
                return null;
            }

            // cannot locate declarations in one pass *the way it's currently implemented*,
            // because the context in EnterSubStmt() doesn't *yet* have child nodes when the context enters.
            // so we need to EnterAmbiguousIdentifier() and evaluate the parent instead - this *might* work.
            var declarationsListener = new DeclarationSymbolsListener(qualifiedName, Accessibility.Implicit, vbComponent.Type, _state.Comments, token);
            
            declarationsListener.NewDeclaration += declarationsListener_NewDeclaration;
            declarationsListener.CreateModuleDeclarations();
            var walker = new ParseTreeWalker();
            walker.Walk(declarationsListener, result.Item1);
            declarationsListener.NewDeclaration -= declarationsListener_NewDeclaration;

            _state.ObsoleteCallContexts = obsoleteCallsListener.Contexts.Select(context => new QualifiedContext(qualifiedName, context));
            _state.ObsoleteLetContexts = obsoleteLetListener.Contexts.Select(context => new QualifiedContext(qualifiedName, context));

            _state.AddTokenStream(vbComponent, result.Item2);

            return result.Item3;
        }

        private void declarationsListener_NewDeclaration(object sender, DeclarationEventArgs e)
        {
             _state.AddDeclaration(e.Declaration);
        }

        private readonly ConcurrentDictionary<VBComponent, CancellationTokenSource> _cancellationTokens =
            new ConcurrentDictionary<VBComponent, CancellationTokenSource>();

        private CancellationToken CancelPreviousTask(VBComponent vbComponent)
        {
            CancellationTokenSource tokenSource;
            if (_cancellationTokens.TryGetValue(vbComponent, out tokenSource))
            {
                try
                {
                    tokenSource.Cancel();
                }
                catch (TaskCanceledException)
                {
                    _state.Status = RubberduckParserState.State.Ready;
                }
            }
            var cancelTokenSource = new CancellationTokenSource();
            _cancellationTokens[vbComponent] = cancelTokenSource;

            return cancelTokenSource.Token;
        }

        private void ResolveReferences(IParseTree tree, CancellationToken token)
        {
            _state.Status = RubberduckParserState.State.Resolving;
            var declarations = _state.AllDeclarations;
            var tasks = _state.UnresolvedDeclarations
                .GroupBy(declaration => declaration.QualifiedSelection.QualifiedName)
                .Where(grouping => grouping.Key.ComponentName != null)
                .Select(grouping => new Task(() =>
                {
                    var resolver = new IdentifierReferenceResolver(grouping.Key, declarations);
                    var listener = new IdentifierReferenceListener(resolver, token);
                    var walker = new ParseTreeWalker();
                    walker.Walk(listener, tree);
                }));
            foreach (var task in tasks)
            {
                task.Start();
            }
        }
    }

    public class ObsoleteCallStatementListener : VBABaseListener
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

    public class ObsoleteLetStatementListener : VBABaseListener
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
