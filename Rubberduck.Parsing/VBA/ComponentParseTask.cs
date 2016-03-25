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
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace Rubberduck.Parsing.VBA
{
    class ComponentParseTask
    {
        private readonly IParseTreeListener[] _listeners;

        private readonly VBComponent _component;
        private readonly QualifiedModuleName _qualifiedName;
        private readonly TokenStreamRewriter _rewriter;
        private readonly IAttributeParser _attributeParser;
        private readonly VBAPreprocessor _preprocessor;

        public event EventHandler<ParseCompletionArgs> ParseCompleted;
        public event EventHandler<ParseFailureArgs> ParseFailure;

        public ComponentParseTask(VBComponent vbComponent, VBAPreprocessor preprocessor, IAttributeParser attributeParser, TokenStreamRewriter rewriter = null)
        {
            _component = vbComponent;
            _listeners = new IParseTreeListener[]
            {
                new ObsoleteCallStatementListener(),
                new ObsoleteLetStatementListener(),
                new EmptyStringLiteralListener(),
                new ArgListWithOneByRefParamListener(),
                new CommentListener(),
            };
            _rewriter = rewriter;
            _qualifiedName = new QualifiedModuleName(vbComponent); 
        }

        public Task ParseAsync(CancellationToken token)
        {
            return new Task(() => ParseInternal(token));
        }

        private void ParseInternal(CancellationToken token)
        {
            try
            {
                var code = RewriteAndPreprocess();
                token.ThrowIfCancellationRequested();

                var stopwatch = Stopwatch.StartNew();
                ITokenStream stream;
                var tree = ParseInternal(code, _listeners, out stream);
                stopwatch.Stop();
                if (tree != null)
                {
                    Debug.Print("IParseTree for component '{0}' acquired in {1}ms (thread {2})", _component.Name, stopwatch.ElapsedMilliseconds, Thread.CurrentThread.ManagedThreadId);
                }

                token.ThrowIfCancellationRequested();

                var attributes = _attributeParser.Parse(_component);
                CommentListener commentListener = _listeners.OfType<CommentListener>().Single();
                var comments = ParseComments(_qualifiedName, commentListener.Comments, commentListener.RemComments);

                token.ThrowIfCancellationRequested();

                var obsoleteCallsListener = _listeners.OfType<ObsoleteCallStatementListener>().Single();
                var obsoleteLetListener = _listeners.OfType<ObsoleteLetStatementListener>().Single();
                var emptyStringLiteralListener = _listeners.OfType<EmptyStringLiteralListener>().Single();
                var argListsWithOneByRefParamListener = _listeners.OfType<ArgListWithOneByRefParamListener>().Single();

                ParseCompleted.Invoke(this, new ParseCompletionArgs
                {
                    Comments = comments,
                    ParseTree = tree,
                    Tokens = stream,
                    Attributes = attributes,
                    ObsoleteCallContexts = obsoleteCallsListener.Contexts.Select(context => new QualifiedContext(_qualifiedName, context)),
                    ObsoleteLetContexts = obsoleteLetListener.Contexts.Select(context => new QualifiedContext(_qualifiedName, context)),
                    EmptyStringLiterals = emptyStringLiteralListener.Contexts.Select(context => new QualifiedContext(_qualifiedName, context)),
                    ArgListsWithOneByRefParam = argListsWithOneByRefParamListener.Contexts.Select(context => new QualifiedContext(_qualifiedName, context)),
                });
            }
            catch (COMException exception)
            {
                Debug.WriteLine("Exception thrown in thread {0}:\n{1}", Thread.CurrentThread.ManagedThreadId, exception);
                ParseFailure.Invoke(this, new ParseFailureArgs
                {
                    Cause = exception
                });
            }
            catch (SyntaxErrorException exception)
            {
                Debug.WriteLine("Exception thrown in thread {0}:\n{1}", Thread.CurrentThread.ManagedThreadId, exception);
                ParseFailure.Invoke(this, new ParseFailureArgs
                {
                    Cause = exception
                });
            }
            catch (OperationCanceledException cancel)
            {
                Debug.WriteLine("Operation was Cancelled", cancel);
                // no results to be used, so no results "returned"
                ParseCompleted.Invoke(this, new ParseCompletionArgs());
            }
        }

        private string RewriteAndPreprocess()
        {
            var code = _rewriter == null ? string.Join(Environment.NewLine, _component.CodeModule.GetSanitizedCode()) : _rewriter.GetText();
            string processed;
            try
            {
                processed = _preprocessor.Execute(code);
            }
            catch (VBAPreprocessorException)
            {
                Debug.WriteLine("Falling back to no preprocessing");
                processed = code;
            }
            return processed;
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


        private IEnumerable<CommentNode> ParseComments(QualifiedModuleName qualifiedName, IEnumerable<VBAParser.CommentContext> comments, IEnumerable<VBAParser.RemCommentContext> remComments)
        {
            var commentNodes = comments.Select(comment => new CommentNode(comment.GetComment(), Tokens.CommentMarker, new QualifiedSelection(qualifiedName, comment.GetSelection())));
            var remCommentNodes = remComments.Select(comment => new CommentNode(comment.GetComment(), Tokens.Rem, new QualifiedSelection(qualifiedName, comment.GetSelection())));
            var allCommentNodes = commentNodes.Union(remCommentNodes);
            return allCommentNodes;
        }

        public void Parse()
        {
            ParseAsync(CancellationToken.None).Wait();
        }

        public class ParseCompletionArgs
        {
            public ITokenStream Tokens { get; internal set; }
            public IParseTree ParseTree { get; internal set; }
            public IEnumerable<CommentNode> Comments { get; internal set; }
            public IEnumerable<QualifiedContext> ObsoleteCallContexts { get; internal set; }
            public IEnumerable<QualifiedContext> ObsoleteLetContexts { get; internal set; }
            public IEnumerable<QualifiedContext> EmptyStringLiterals { get; internal set; }
            public IEnumerable<QualifiedContext> ArgListsWithOneByRefParam { get; internal set; }
            public IEnumerable<Declaration> Declarations { get; internal set; }
            public IDictionary<Tuple<string, DeclarationType>, Attributes> Attributes { get; internal set; }
        }

        public class ParseFailureArgs
        {
            public Exception Cause { get; internal set; }
        }
    }

    #region Listener classes
    class ObsoleteCallStatementListener : VBABaseListener
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

    class ObsoleteLetStatementListener : VBABaseListener
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

    class EmptyStringLiteralListener : VBABaseListener
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

    class ArgListWithOneByRefParamListener : VBABaseListener
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

    class CommentListener : VBABaseListener
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

    #endregion
}
