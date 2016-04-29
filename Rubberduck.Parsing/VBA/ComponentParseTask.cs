using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Annotations;
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

namespace Rubberduck.Parsing.VBA
{
    class ComponentParseTask
    {
        private readonly VBComponent _component;
        private readonly QualifiedModuleName _qualifiedName;
        private readonly TokenStreamRewriter _rewriter;
        private readonly IAttributeParser _attributeParser;
        private readonly VBAPreprocessor _preprocessor;

        public event EventHandler<ParseCompletionArgs> ParseCompleted;
        public event EventHandler<ParseFailureArgs> ParseFailure;

        public ComponentParseTask(VBComponent vbComponent, VBAPreprocessor preprocessor, IAttributeParser attributeParser, TokenStreamRewriter rewriter = null)
        {
            _attributeParser = attributeParser;
            _preprocessor = preprocessor;
            _component = vbComponent;
            _rewriter = rewriter;
            _qualifiedName = new QualifiedModuleName(vbComponent); 
        }
        
        public void Start(CancellationToken token)
        {
            try
            {
                var code = RewriteAndPreprocess();
                token.ThrowIfCancellationRequested();

                var attributes = _attributeParser.Parse(_component);

                token.ThrowIfCancellationRequested();

                // temporal coupling... comments must be acquired before we walk the parse tree for declarations
                // otherwise none of the annotations get associated to their respective Declaration
                var commentListener = new CommentListener();
                var annotationListener = new AnnotationListener(new VBAParserAnnotationFactory(), _qualifiedName);

                var stopwatch = Stopwatch.StartNew();
                ITokenStream stream;
                var tree = ParseInternal(code, new IParseTreeListener[]{ commentListener, annotationListener }, out stream);
                stopwatch.Stop();
                if (tree != null)
                {
                    Debug.Print("IParseTree for component '{0}' acquired in {1}ms (thread {2})", _component.Name, stopwatch.ElapsedMilliseconds, Thread.CurrentThread.ManagedThreadId);
                }

                var comments = QualifyAndUnionComments(_qualifiedName, commentListener.Comments, commentListener.RemComments);
                token.ThrowIfCancellationRequested();

                ParseCompleted.Invoke(this, new ParseCompletionArgs
                {
                    ParseTree = tree,
                    Tokens = stream,
                    Attributes = attributes,
                    Comments = comments,
                    Annotations = annotationListener.Annotations
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
                //ParseCompleted.Invoke(this, new ParseCompletionArgs());
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

        private static IParseTree ParseInternal(string code, IParseTreeListener[] listeners, out ITokenStream outStream)
        {
            var stream = new AntlrInputStream(code);
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAParser(tokens);

            parser.AddErrorListener(new ExceptionErrorListener());
            foreach (var l in listeners)
            {
                parser.AddParseListener(l);
            }

            outStream = tokens;
            return parser.startRule();
        }

        private IEnumerable<CommentNode> QualifyAndUnionComments(QualifiedModuleName qualifiedName, IEnumerable<VBAParser.CommentContext> comments, IEnumerable<VBAParser.RemCommentContext> remComments)
        {
            var commentNodes = comments.Select(comment => new CommentNode(comment.GetComment(), Tokens.CommentMarker, new QualifiedSelection(qualifiedName, comment.GetSelection())));
            var remCommentNodes = remComments.Select(comment => new CommentNode(comment.GetComment(), Tokens.Rem, new QualifiedSelection(qualifiedName, comment.GetSelection())));
            var allCommentNodes = commentNodes.Union(remCommentNodes);
            return allCommentNodes;
        }
        
        public class ParseCompletionArgs
        {
            public ITokenStream Tokens { get; internal set; }
            public IParseTree ParseTree { get; internal set; }
            public IDictionary<Tuple<string, DeclarationType>, Attributes> Attributes { get; internal set; }
            public IEnumerable<CommentNode> Comments { get; internal set; }
            public IEnumerable<IAnnotation> Annotations { get; internal set; }
        }

        public class ParseFailureArgs
        {
            public Exception Cause { get; internal set; }
        }

        private class CommentListener : VBAParserBaseListener
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
