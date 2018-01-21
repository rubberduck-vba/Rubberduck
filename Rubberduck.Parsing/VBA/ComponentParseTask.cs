using Antlr4.Runtime;
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
using Rubberduck.Parsing.Symbols.ParsingExceptions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA
{
    class ComponentParseTask
    {
        private readonly QualifiedModuleName _module;
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
            _module = module;
            _rewriter = rewriter;
            _parser = new VBAModuleParser();
        }
        
        public void Start(CancellationToken cancellationToken)
        {
            try
            {
                Logger.Trace($"Starting ParseTaskID {_taskId} on thread {Thread.CurrentThread.ManagedThreadId}.");

                var tokenStream = RewriteAndPreprocess(cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();  

                // temporal coupling... comments must be acquired before we walk the parse tree for declarations
                // otherwise none of the annotations get associated to their respective Declaration
                var commentListener = new CommentListener();
                var annotationListener = new AnnotationListener(new VBAParserAnnotationFactory(), _module);

                var stopwatch = Stopwatch.StartNew();
                var codePaneParseResults = ParseInternal(_module.ComponentName, tokenStream, new IParseTreeListener[]{ commentListener, annotationListener });
                stopwatch.Stop();
                cancellationToken.ThrowIfCancellationRequested();

                var comments = QualifyAndUnionComments(_module, commentListener.Comments, commentListener.RemComments);
                cancellationToken.ThrowIfCancellationRequested();

                var attributesPassParseResults = RunAttributesPass(cancellationToken);
                var rewriter = new MemberAttributesRewriter(_exporter, _module.Component.CodeModule, new TokenStreamRewriter(attributesPassParseResults.tokenStream ?? tokenStream));

                var completedHandler = ParseCompleted;
                if (completedHandler != null && !cancellationToken.IsCancellationRequested)
                    completedHandler.Invoke(this, new ParseCompletionArgs
                    {
                        ParseTree = codePaneParseResults.tree,
                        AttributesTree = attributesPassParseResults.tree,
                        Tokens = codePaneParseResults.tokenStream,
                        AttributesRewriter = rewriter,
                        Attributes = attributesPassParseResults.attributes,
                        Comments = comments,
                        Annotations = annotationListener.Annotations
                    });
            }
            catch (COMException exception)
            {
                Logger.Error(exception, $"COM Exception thrown in thread {Thread.CurrentThread.ManagedThreadId} while parsing module {_module.ComponentName}, ParseTaskID {_taskId}.");
                var failedHandler = ParseFailure;
                failedHandler?.Invoke(this, new ParseFailureArgs
                {
                    Cause = exception
                });
            }
            catch (PreprocessorSyntaxErrorException syntaxErrorException)
            {
                var parsePassText = syntaxErrorException.ParsePass == ParsePass.CodePanePass
                    ? "code pane"
                    : "exported";
                Logger.Error($"Syntax error while preprocessing; offending token '{syntaxErrorException.OffendingSymbol.Text}' at line {syntaxErrorException.LineNumber}, column {syntaxErrorException.Position} in the {parsePassText} version of module {_module.ComponentName}.");
                Logger.Debug(syntaxErrorException, $"SyntaxErrorException thrown in thread {Thread.CurrentThread.ManagedThreadId}, ParseTaskID {_taskId}.");

                ReportException(syntaxErrorException);
            }
            catch (ParsePassSyntaxErrorException syntaxErrorException)
            {
                var parsePassText = syntaxErrorException.ParsePass == ParsePass.CodePanePass
                    ? "code pane"
                    : "exported";
                Logger.Error($"Syntax error; offending token '{syntaxErrorException.OffendingSymbol.Text}' at line {syntaxErrorException.LineNumber}, column {syntaxErrorException.Position} in the {parsePassText} version of module {_module.ComponentName}.");
                Logger.Debug(syntaxErrorException, $"SyntaxErrorException thrown in thread {Thread.CurrentThread.ManagedThreadId}, ParseTaskID {_taskId}.");

                ReportException(syntaxErrorException);
            }
            catch (SyntaxErrorException syntaxErrorException)
            {
                Logger.Error($"Syntax error; offending token '{syntaxErrorException.OffendingSymbol.Text}' at line {syntaxErrorException.LineNumber}, column {syntaxErrorException.Position} in module {_module.ComponentName}.");
                Logger.Debug(syntaxErrorException, $"SyntaxErrorException thrown in thread {Thread.CurrentThread.ManagedThreadId}, ParseTaskID {_taskId}.");

                ReportException(syntaxErrorException);
            }
            catch (OperationCanceledException exception)
            {
                //We report this, so that the calling code knows that the operation actually has been cancelled.
                ReportException(exception);
            }
            catch (Exception exception)
            {
                Logger.Error(exception, $" Unexpected exception thrown in thread {Thread.CurrentThread.ManagedThreadId} while parsing module {_module.ComponentName}, ParseTaskID {_taskId}.");

                ReportException(exception);
            }
        }

        private void ReportException(Exception exception)
        {
            var failedHandler = ParseFailure;
            failedHandler?.Invoke(this, new ParseFailureArgs
            {
                Cause = exception
            });
        }

        private (IParseTree tree, ITokenStream tokenStream, IDictionary<Tuple<string, DeclarationType>, Attributes> attributes) RunAttributesPass(CancellationToken cancellationToken)
        {
            Logger.Trace($"ParseTaskID {_taskId} begins attributes pass.");
            var attributesParseResults = _attributeParser.Parse(_module, cancellationToken);
            Logger.Trace($"ParseTaskID {_taskId} finished attributes pass.");
            return attributesParseResults;
        }

        private static string GetCode(IVBComponent component)
        {
            string codeLines;
            using (var module = component.CodeModule)
            {
                var lines = module.CountOfLines;
                if (lines == 0)
                {
                    return string.Empty;
                }

                codeLines = module.GetLines(1, lines);
            }
            var code = string.Concat(codeLines);

            return code;
        }

        private CommonTokenStream RewriteAndPreprocess(CancellationToken cancellationToken)
        {
            var code = _rewriter?.GetText() ?? string.Join(Environment.NewLine, GetCode(_module.Component));
            var tokenStreamProvider = new SimpleVBAModuleTokenStreamProvider();
            var tokens = tokenStreamProvider.Tokens(code);
            _preprocessor.PreprocessTokenStream(_module.Name, tokens, new PreprocessorExceptionErrorListener(_module.ComponentName, ParsePass.CodePanePass), cancellationToken);
            return tokens;
        }

        private (IParseTree tree, ITokenStream tokenStream) ParseInternal(string moduleName, CommonTokenStream tokenStream, IParseTreeListener[] listeners)
        {
            //var errorNotifier = new SyntaxErrorNotificationListener();
            //errorNotifier.OnSyntaxError += ParserSyntaxError;
            return _parser.Parse(moduleName, tokenStream, listeners, new MainParseExceptionErrorListener(moduleName, ParsePass.CodePanePass));
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
