using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.Symbols.ParsingExceptions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.Parsing.VBA.Parsing
{
    class ComponentParseTask
    {
        private readonly QualifiedModuleName _module;
        private readonly TokenStreamRewriter _rewriter;
        private readonly ISourceCodeProvider _codePaneSourceCodeProvider;
        private readonly ISourceCodeProvider _attributesSourceCodeProvider;
        private readonly IStringParser _parser;
        private readonly IModuleRewriterFactory _moduleRewriterFactory;

        public event EventHandler<ParseCompletionArgs> ParseCompleted;
        public event EventHandler<ParseFailureArgs> ParseFailure;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly Guid _taskId;

        public ComponentParseTask(QualifiedModuleName module, ISourceCodeProvider codePaneSourceCodeProvider, ISourceCodeProvider attributesSourceCodeProvider, IStringParser parser, IModuleRewriterFactory moduleRewriterFactory, TokenStreamRewriter rewriter = null)
        {
            _taskId = Guid.NewGuid();

            _moduleRewriterFactory = moduleRewriterFactory;
            _codePaneSourceCodeProvider = codePaneSourceCodeProvider;
            _attributesSourceCodeProvider = attributesSourceCodeProvider;
            _module = module;
            _rewriter = rewriter;
            _parser = parser;
        }
        
        public void Start(CancellationToken cancellationToken)
        {
            try
            {
                Logger.Trace($"Starting ParseTaskID {_taskId} on thread {Thread.CurrentThread.ManagedThreadId}.");

                cancellationToken.ThrowIfCancellationRequested();  

                var (codePaneParseTree, codePaneTokenStream) = CodePanePassResults(_module, cancellationToken, _rewriter);
                var codePaneRewriter = _moduleRewriterFactory.CodePaneRewriter(_module, codePaneTokenStream);
                cancellationToken.ThrowIfCancellationRequested();

                // temporal coupling... comments must be acquired before we walk the parse tree for declarations
                // otherwise none of the annotations get associated to their respective Declaration
                var (comments, annotations) = CommentsAndAnnotations(_module, codePaneParseTree);
                cancellationToken.ThrowIfCancellationRequested();

                var (attributesParseTree, attributesTokenStream) = AttributesPassResults(_module, cancellationToken);
                var attributesRewriter = _moduleRewriterFactory.AttributesRewriter(_module, attributesTokenStream ?? codePaneTokenStream);
                cancellationToken.ThrowIfCancellationRequested();

                var attributes = Attributes(_module, attributesParseTree);
                cancellationToken.ThrowIfCancellationRequested();

                var completedHandler = ParseCompleted;
                if (completedHandler != null && !cancellationToken.IsCancellationRequested)
                    completedHandler.Invoke(this, new ParseCompletionArgs
                    {
                        ParseTree = codePaneParseTree,
                        AttributesTree = attributesParseTree,
                        CodePaneRewriter = codePaneRewriter,
                        AttributesRewriter = attributesRewriter,
                        Attributes = attributes,
                        Comments = comments,
                        Annotations = annotations
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
                var parsePassText = syntaxErrorException.CodeKind == CodeKind.CodePaneCode
                    ? "code pane"
                    : "exported";
                Logger.Error($"Syntax error while preprocessing; offending token '{syntaxErrorException.OffendingSymbol.Text}' at line {syntaxErrorException.LineNumber}, column {syntaxErrorException.Position} in the {parsePassText} version of module {_module.ComponentName}.");
                Logger.Debug(syntaxErrorException, $"SyntaxErrorException thrown in thread {Thread.CurrentThread.ManagedThreadId}, ParseTaskID {_taskId}.");

                ReportException(syntaxErrorException);
            }
            catch (ParsePassSyntaxErrorException syntaxErrorException)
            {
                var parsePassText = syntaxErrorException.CodeKind == CodeKind.CodePaneCode
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

        private (IEnumerable<CommentNode> Comments, IEnumerable<IAnnotation> Annotations) CommentsAndAnnotations(QualifiedModuleName module, IParseTree tree)
        {
            var commentListener = new CommentListener();
            var annotationListener = new AnnotationListener(new VBAParserAnnotationFactory(), _module);
            var combinedListener = new CombinedParseTreeListener(new IParseTreeListener[] {commentListener, annotationListener});
            ParseTreeWalker.Default.Walk(combinedListener, tree);
            var comments = QualifyAndUnionComments(module, commentListener.Comments, commentListener.RemComments);
            var annotations = annotationListener.Annotations;
            return (comments, annotations);
        }

        private IDictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes> Attributes(QualifiedModuleName module, IParseTree tree)
        {
            var type = module.ComponentType == ComponentType.StandardModule
                ? DeclarationType.ProceduralModule
                : DeclarationType.ClassModule;
            var attributesListener = new AttributeListener((module.ComponentName, type));
            ParseTreeWalker.Default.Walk(attributesListener, tree);
            return attributesListener.Attributes;
        }

        private void ReportException(Exception exception)
        {
            var failedHandler = ParseFailure;
            failedHandler?.Invoke(this, new ParseFailureArgs
            {
                Cause = exception
            });
        }

        private (IParseTree tree, ITokenStream tokenStream) AttributesPassResults(QualifiedModuleName module, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            Logger.Trace($"ParseTaskID {_taskId} begins attributes pass.");
            var code = _attributesSourceCodeProvider.SourceCode(module);
            token.ThrowIfCancellationRequested();
            var attributesParseResults = _parser.Parse(module.ComponentName, module.ProjectId, code, token, CodeKind.AttributesCode);
            Logger.Trace($"ParseTaskID {_taskId} finished attributes pass.");
            return attributesParseResults;
        }

        private static string GetCode(ICodeModule codeModule)
        {
            var lines = codeModule.CountOfLines;
            if (lines == 0)
            {
                return string.Empty;
            }

            var codeLines = codeModule.GetLines(1, lines);
            var code = string.Concat(codeLines);

            return code;
        }

        private (IParseTree tree, ITokenStream tokenStream) CodePanePassResults(QualifiedModuleName module, CancellationToken token, TokenStreamRewriter rewriter = null)
        {
            token.ThrowIfCancellationRequested();
            var code = rewriter?.GetText() ?? _codePaneSourceCodeProvider.SourceCode(module);
            token.ThrowIfCancellationRequested();
            return _parser.Parse(module.ComponentName, module.ProjectId, code, token, CodeKind.CodePaneCode);
        }

        private IEnumerable<CommentNode> QualifyAndUnionComments(QualifiedModuleName qualifiedName, IEnumerable<VBAParser.CommentContext> comments, IEnumerable<VBAParser.RemCommentContext> remComments)
        {
            var commentNodes = comments.Select(comment => new CommentNode(CommentExtensions.GetComment((VBAParser.CommentContext) comment), Tokens.CommentMarker, new QualifiedSelection(qualifiedName, ParserRuleContextExtensions.GetSelection(comment))));
            var remCommentNodes = remComments.Select(comment => new CommentNode(comment.GetComment(), Tokens.Rem, new QualifiedSelection(qualifiedName, comment.GetSelection())));
            var allCommentNodes = commentNodes.Union(remCommentNodes);
            return allCommentNodes;
        }
        
        public class ParseCompletionArgs
        {
            public IModuleRewriter CodePaneRewriter { get; internal set; }
            public IModuleRewriter AttributesRewriter { get; internal set; }
            public IParseTree ParseTree { get; internal set; }
            public IParseTree AttributesTree { get; internal set; }
            public IDictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes> Attributes { get; internal set; }
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
