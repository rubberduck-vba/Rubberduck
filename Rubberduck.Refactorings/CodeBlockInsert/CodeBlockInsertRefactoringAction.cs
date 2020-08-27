using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.CodeBlockInsert
{
    /// <summary>
    /// Inserts new content into a ModuleDeclarationSection of CodeSection
    /// based on the presence or absence of existing CodeSection content.  
    /// </summary>
    /// <remarks>If there is an existing member <c>Declaration</c>, then the
    /// new content will be inserted above the existing member.
    /// </remarks>
    public class CodeBlockInsertRefactoringAction : CodeOnlyRefactoringActionBase<CodeBlockInsertModel>
    {
        private static string _doubleSpace = $"{Environment.NewLine}{Environment.NewLine}";

        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ICodeBuilder _codeBuilder;

        public CodeBlockInsertRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager, ICodeBuilder codeBuilder)
            : base(rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _codeBuilder = codeBuilder;
        }

        public override void Refactor(CodeBlockInsertModel model, IRewriteSession rewriteSession)
        {
            var newDeclarationSectionBlock = BuildBlock(model, NewContentType.TypeDeclarationBlock, NewContentType.DeclarationBlock);

            var newCodeSectionBlock = BuildBlock(model, NewContentType.CodeSectionBlock);

            var aggregatedBlocks = new List<string>();
            if (!string.IsNullOrEmpty(newDeclarationSectionBlock))
            {
                aggregatedBlocks.Add(newDeclarationSectionBlock);
            }

            if (!string.IsNullOrEmpty(newCodeSectionBlock))
            {
                aggregatedBlocks.Add(newCodeSectionBlock);
            }

            if (aggregatedBlocks.Count == 0)
            {
                return;
            }

            var rewriter = rewriteSession.CheckOutModuleRewriter(model.QualifiedModuleName);

            var allNewContent = string.Join(_doubleSpace, aggregatedBlocks);

            var comments = string.Join(_doubleSpace, model.NewContent[NewContentType.PostContentMessage]);
            if (!string.IsNullOrEmpty(comments) && model.IncludeComments)
            {
                allNewContent = $"{allNewContent}{Environment.NewLine}{comments}";
            }

            InsertBlock(allNewContent, model.NewContentInsertionIndex, rewriter);
        }

        private static void InsertBlock(string content, int? insertionIndex, IModuleRewriter rewriter)
        {
            if (string.IsNullOrEmpty(content))
            {
                return;
            }

            if (insertionIndex.HasValue)
            {
                rewriter.InsertBefore(insertionIndex.Value, $"{content}{_doubleSpace}");
                return;
            }
            rewriter.InsertBefore(rewriter.TokenStream.Size - 1, $"{_doubleSpace}{content}");
        }

        private string BuildBlock(CodeBlockInsertModel model, params NewContentType[] newContentTypes)
        {
            var block = string.Empty;
            foreach (var newContentType in newContentTypes)
            {
                var newContent = string.Join(_doubleSpace,
                    (model.NewContent[newContentType]));
                if (!string.IsNullOrEmpty(newContent))
                {
                    block = string.IsNullOrEmpty(block)
                        ? newContent
                        : $"{block}{_doubleSpace}{newContent}";
                }
            }
            return LimitNewLines(block.Trim(), model.NewLineLimit);
        }

        private static string LimitNewLines(string content, int maxConsecutiveNewlines = 2)
        {
            var target = string.Concat(Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewlines + 1).ToList());
            var replacement = string.Concat(Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewlines).ToList());
            var guard = 0;
            var maxAttempts = 100;
            while (++guard < maxAttempts && content.Contains(target))
            {
                content = content.Replace(target, replacement);
            }

            if (guard >= maxAttempts)
            {
                throw new FormatException($"Unable to limit consecutive '{Environment.NewLine}' strings to {maxConsecutiveNewlines}");
            }
            return content;
        }

    }
}
