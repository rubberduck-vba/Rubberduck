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
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.Parsing.ParsingExceptions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.Parsing.VBA.Parsing
{
    public class ModuleParser : IModuleParser
    {
        private readonly ISourceCodeProvider _codePaneSourceCodeProvider;
        private readonly ISourceCodeProvider _attributesSourceCodeProvider;
        private readonly IStringParser _parser;
        private readonly IAnnotationFactory _annotationFactory;

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public ModuleParser(ISourceCodeProvider codePaneSourceCodeProvider, ISourceCodeProvider attributesSourceCodeProvider, IStringParser parser, IAnnotationFactory annotationFactory)
        {
            _codePaneSourceCodeProvider = codePaneSourceCodeProvider;
            _attributesSourceCodeProvider = attributesSourceCodeProvider;
            _parser = parser;
            _annotationFactory = annotationFactory;
        }

        public ModuleParseResults Parse(QualifiedModuleName module, CancellationToken cancellationToken, TokenStreamRewriter rewriter = null)
        {
            var taskId = Guid.NewGuid();
            try
            {
                return ParseInternal(module, cancellationToken, rewriter, taskId);
            }
            catch (COMException exception)
            {
                Logger.Error(exception, $"COM Exception thrown in thread {Thread.CurrentThread.ManagedThreadId} while parsing module {module.ComponentName}, ParseTaskID {taskId}.");
                throw;
            }
            catch (PreprocessorSyntaxErrorException syntaxErrorException)
            {
                Logger.Error($"Syntax error while preprocessing; offending token '{syntaxErrorException.OffendingSymbol.Text}' at line {syntaxErrorException.LineNumber}, column {syntaxErrorException.Position} in the {syntaxErrorException.CodeKind} version of module {module.ComponentName}.");
                Logger.Debug(syntaxErrorException, $"SyntaxErrorException thrown in thread {Thread.CurrentThread.ManagedThreadId}, ParseTaskID {taskId}.");
                throw;
            }
            catch (ParsePassSyntaxErrorException syntaxErrorException)
            {
                Logger.Error($"Syntax error; offending token '{syntaxErrorException.OffendingSymbol.Text}' at line {syntaxErrorException.LineNumber}, column {syntaxErrorException.Position} in the {syntaxErrorException.CodeKind} version of module {module.ComponentName}.");
                Logger.Debug(syntaxErrorException, $"SyntaxErrorException thrown in thread {Thread.CurrentThread.ManagedThreadId}, ParseTaskID {taskId}.");
                throw;
            }
            catch (SyntaxErrorException syntaxErrorException)
            {
                Logger.Error($"Syntax error; offending token '{syntaxErrorException.OffendingSymbol.Text}' at line {syntaxErrorException.LineNumber}, column {syntaxErrorException.Position} in module {module.ComponentName}.");
                Logger.Debug(syntaxErrorException, $"SyntaxErrorException thrown in thread {Thread.CurrentThread.ManagedThreadId}, ParseTaskID {taskId}.");
                throw;
            }
            catch (OperationCanceledException exception)
            {
                //We rethrow this, so that the calling code knows that the operation actually has been cancelled.
                //No need to log it.
                throw;
            }
            catch (Exception exception)
            {
                Logger.Error(exception, $" Unexpected exception thrown in thread {Thread.CurrentThread.ManagedThreadId} while parsing module {module.ComponentName}, ParseTaskID {taskId}.");
                throw;
            }
        }

        private ModuleParseResults ParseInternal(QualifiedModuleName module, CancellationToken cancellationToken, TokenStreamRewriter rewriter, Guid taskId)
        {
            Logger.Trace($"Starting ParseTaskID {taskId} on thread {Thread.CurrentThread.ManagedThreadId}.");

            cancellationToken.ThrowIfCancellationRequested();

            Logger.Trace($"ParseTaskID {taskId} begins code pane pass.");
            var (codePaneParseTree, codePaneTokenStream, logicalLines) = CodePanePassResults(module, cancellationToken, rewriter);
            Logger.Trace($"ParseTaskID {taskId} finished code pane pass.");
            cancellationToken.ThrowIfCancellationRequested();

            // temporal coupling... comments must be acquired before we walk the parse tree for declarations
            // otherwise none of the annotations get associated to their respective Declaration
            Logger.Trace($"ParseTaskID {taskId} begins extracting comments and annotations.");
            var (comments, annotations) = CommentsAndAnnotations(module, codePaneParseTree);
            Logger.Trace($"ParseTaskID {taskId} finished extracting comments and annotations.");
            cancellationToken.ThrowIfCancellationRequested();

            Logger.Trace($"ParseTaskID {taskId} begins attributes pass.");
            var (attributesParseTree, attributesTokenStream) = AttributesPassResults(module, cancellationToken);
            Logger.Trace($"ParseTaskID {taskId} finished attributes pass.");
            cancellationToken.ThrowIfCancellationRequested();

            Logger.Trace($"ParseTaskID {taskId} begins extracting attributes.");
            var (attributes, membersAllowingAttributes) = Attributes(module, attributesParseTree);
            Logger.Trace($"ParseTaskID {taskId} finished extracting attributes.");
            cancellationToken.ThrowIfCancellationRequested();

            return new ModuleParseResults(
                codePaneParseTree,
                attributesParseTree,
                comments,
                annotations,
                logicalLines,
                attributes,
                membersAllowingAttributes,
                codePaneTokenStream,
                attributesTokenStream
            );            
        }

        private (IEnumerable<CommentNode> Comments, IEnumerable<IParseTreeAnnotation> Annotations) CommentsAndAnnotations(QualifiedModuleName module, IParseTree tree)
        {
            var commentListener = new CommentListener();
            var annotationListener = new AnnotationListener(_annotationFactory, module);
            var combinedListener = new CombinedParseTreeListener(new IParseTreeListener[] {commentListener, annotationListener});
            ParseTreeWalker.Default.Walk(combinedListener, tree);
            var comments = QualifyAndUnionComments(module, commentListener.Comments, commentListener.RemComments);
            var annotations = annotationListener.Annotations;
            return (comments, annotations);
        }

        private IEnumerable<CommentNode> QualifyAndUnionComments(QualifiedModuleName qualifiedName, IEnumerable<VBAParser.CommentContext> comments, IEnumerable<VBAParser.RemCommentContext> remComments)
        {
            var commentNodes = comments.Select(comment => new CommentNode(((VBAParser.CommentContext)comment).GetComment(), Tokens.CommentMarker, new QualifiedSelection(qualifiedName, comment.GetSelection())));
            var remCommentNodes = remComments.Select(comment => new CommentNode(comment.GetComment(), Tokens.Rem, new QualifiedSelection(qualifiedName, comment.GetSelection())));
            var allCommentNodes = commentNodes.Union(remCommentNodes);
            return allCommentNodes;
        }

        private (IDictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes> attributes,
            IDictionary<(string scopeIdentifier, DeclarationType scopeType), ParserRuleContext> membersAllowingAttributes) 
            Attributes(QualifiedModuleName module, IParseTree tree)
        {
            var type = module.ComponentType == ComponentType.StandardModule
                ? DeclarationType.ProceduralModule
                : DeclarationType.ClassModule;
            var attributesListener = new AttributeListener((module.ComponentName, type));
            ParseTreeWalker.Default.Walk(attributesListener, tree);
            return (attributesListener.Attributes, attributesListener.MembersAllowingAttributes);
        }

        private (IParseTree tree, ITokenStream tokenStream) AttributesPassResults(QualifiedModuleName module, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            var code = _attributesSourceCodeProvider.SourceCode(module);
            token.ThrowIfCancellationRequested();
            var attributesParseResults = _parser.Parse(module.ComponentName, module.ProjectId, code, token, CodeKind.AttributesCode);
            return attributesParseResults;
        }

        private (IParseTree tree, ITokenStream tokenStream, LogicalLineStore logicalLines) CodePanePassResults(QualifiedModuleName module, CancellationToken token, TokenStreamRewriter rewriter = null)
        {
            token.ThrowIfCancellationRequested();
            var code = rewriter?.GetText() ?? _codePaneSourceCodeProvider.SourceCode(module);
            var logicalLines = LogicalLines(code);
            token.ThrowIfCancellationRequested();
            var (tree, tokenStream) = _parser.Parse(module.ComponentName, module.ProjectId, code, token, CodeKind.CodePaneCode);
            return (tree, tokenStream, logicalLines);
        }

        private LogicalLineStore LogicalLines(string code)
        {
            var lines = code.Split(new []{Environment.NewLine}, StringSplitOptions.None);
            var logicalLineEnds = lines
                .Select((line, index) => (line, index))
                .Where(tpl => !tpl.line.TrimEnd().EndsWith(" _")) //Not line-continued
                .Select(tpl => tpl.index + 1); //VBA lines are 1-based.
            return new LogicalLineStore(logicalLineEnds);
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
