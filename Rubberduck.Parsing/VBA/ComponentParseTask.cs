﻿using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA
{
    class ComponentParseTask
    {
        private readonly IVBComponent _component;
        private readonly QualifiedModuleName _qualifiedName;
        private readonly TokenStreamRewriter _rewriter;
        private readonly IAttributeParser _attributeParser;
        private readonly IModuleExporter _exporter;
        private readonly IVBAPreprocessor _preprocessor;
        private readonly VBAModuleParser _parser;

        public event EventHandler<ParseCompletionArgs> ParseCompleted;
        public event EventHandler<ParseFailureArgs> ParseFailure;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly Guid _taskId;

        public ComponentParseTask(QualifiedModuleName module, IVBAPreprocessor preprocessor, IAttributeParser attributeParser, IModuleExporter exporter, TokenStreamRewriter rewriter = null)
        {
            _taskId = Guid.NewGuid();

            _attributeParser = attributeParser;
            _exporter = exporter;
            _preprocessor = preprocessor;
            _component = module.Component;
            _rewriter = rewriter;
            _qualifiedName = module;
            _parser = new VBAModuleParser();
        }
        
        public void Start(CancellationToken token)
        {
            try
            {
                Logger.Trace($"Starting ParseTaskID {_taskId} on thread {Thread.CurrentThread.ManagedThreadId}.");

                var tokenStream = RewriteAndPreprocess(token);
                token.ThrowIfCancellationRequested();

                IParseTree attributesTree;
                IDictionary<Tuple<string, DeclarationType>, Attributes> attributes;
                var attributesTokenStream = RunAttributesPass(token, out attributesTree, out attributes);

                var rewriter = new MemberAttributesRewriter(_exporter, _component.CodeModule, new TokenStreamRewriter(attributesTokenStream ?? tokenStream));

                // temporal coupling... comments must be acquired before we walk the parse tree for declarations
                // otherwise none of the annotations get associated to their respective Declaration
                var commentListener = new CommentListener();
                var annotationListener = new AnnotationListener(new VBAParserAnnotationFactory(), _qualifiedName);

                var stopwatch = Stopwatch.StartNew();
                ITokenStream stream;
                var tree = ParseInternal(_component.Name, tokenStream, new IParseTreeListener[]{ commentListener, annotationListener }, out stream);
                stopwatch.Stop();
                token.ThrowIfCancellationRequested();

                var comments = QualifyAndUnionComments(_qualifiedName, commentListener.Comments, commentListener.RemComments);
                token.ThrowIfCancellationRequested();

                var completedHandler = ParseCompleted;
                if (completedHandler != null && !token.IsCancellationRequested)
                    completedHandler.Invoke(this, new ParseCompletionArgs
                    {
                        ParseTree = tree,
                        AttributesTree = attributesTree,
                        Tokens = stream,
                        AttributesRewriter = rewriter,
                        Attributes = attributes,
                        Comments = comments,
                        Annotations = annotationListener.Annotations
                    });
            }
            catch (COMException exception)
            {
                Logger.Error(exception, $"Exception thrown in thread {Thread.CurrentThread.ManagedThreadId} while parsing module {_qualifiedName.Name}, ParseTaskID {_taskId}.");
                var failedHandler = ParseFailure;
                failedHandler?.Invoke(this, new ParseFailureArgs
                {
                    Cause = exception
                });
            }
            catch (SyntaxErrorException exception)
            {
                Logger.Warn($"Syntax error; offending token '{exception.OffendingSymbol.Text}' at line {exception.LineNumber}, column {exception.Position} in module {_qualifiedName}.");
                Logger.Error(exception, $"Exception thrown in thread {Thread.CurrentThread.ManagedThreadId}, ParseTaskID {_taskId}.");
                var failedHandler = ParseFailure;
                failedHandler?.Invoke(this, new ParseFailureArgs
                {
                    Cause = exception
                });
            }
            catch (OperationCanceledException exception)
            {
                //We return this, so that the calling code knows that the operation actually has been cancelled.
                var failedHandler = ParseFailure;
                failedHandler?.Invoke(this, new ParseFailureArgs
                {
                    Cause = exception
                });
            }
            catch (Exception exception)
            {
                Logger.Error(exception, $"Exception thrown in thread {Thread.CurrentThread.ManagedThreadId} while parsing module {_qualifiedName.Name}, ParseTaskID {_taskId}.");
                var failedHandler = ParseFailure;
                failedHandler?.Invoke(this, new ParseFailureArgs
                {
                    Cause = exception
                });
            }
        }

        private ITokenStream RunAttributesPass(CancellationToken token, out IParseTree attributesTree,
            out IDictionary<Tuple<string, DeclarationType>, Attributes> attributes)
        {
            Logger.Trace($"ParseTaskID {_taskId} begins attributes pass.");
            ITokenStream attributesTokenStream;
            attributes = _attributeParser.Parse(_component, token, out attributesTokenStream, out attributesTree);
            Logger.Trace($"ParseTaskID {_taskId} finished attributes pass.");
            return attributesTokenStream;
        }

        private static string GetCode(ICodeModule module)
        {
            var lines = module.CountOfLines;
            if (lines == 0)
            {
                return string.Empty;
            }

            var codeLines = module.GetLines(1, lines);
            var code = string.Concat(codeLines);

            return code;
        }

        private CommonTokenStream RewriteAndPreprocess(CancellationToken token)
        {
            var code = _rewriter == null ? string.Join(Environment.NewLine, GetCode(_component.CodeModule)) : _rewriter.GetText();
            var tokenStreamProvider = new SimpleVBAModuleTokenStreamProvider();
            var tokens = tokenStreamProvider.Tokens(code);
            _preprocessor.PreprocessTokenStream(_component.Name, tokens, token);
            return tokens;
        }

        private IParseTree ParseInternal(string moduleName, CommonTokenStream tokenStream, IParseTreeListener[] listeners, out ITokenStream outStream)
        {
            //var errorNotifier = new SyntaxErrorNotificationListener();
            //errorNotifier.OnSyntaxError += ParserSyntaxError;
            return _parser.Parse(moduleName, tokenStream, listeners, new ExceptionErrorListener(), out outStream);
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
            public IModuleRewriter AttributesRewriter { get; internal set; }
            public IParseTree ParseTree { get; internal set; }
            public IParseTree AttributesTree { get; internal set; }
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
            public IEnumerable<VBAParser.RemCommentContext> RemComments => _remComments;

            private readonly IList<VBAParser.CommentContext> _comments = new List<VBAParser.CommentContext>();
            public IEnumerable<VBAParser.CommentContext> Comments => _comments;

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
