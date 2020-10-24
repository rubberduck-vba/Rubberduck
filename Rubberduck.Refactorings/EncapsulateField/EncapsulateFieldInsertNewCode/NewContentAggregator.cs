using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Resources;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public enum NewContentType
    {
        UserDefinedTypeDeclaration,
        DeclarationBlock,
        CodeSectionBlock,
    }

    public interface INewContentAggregatorFactory
    {
        INewContentAggregator Create();
    }

    public interface INewContentAggregator
    {
        /// <summary>
        /// Allows gouping content blocks by <c>NewContentType</c>.
        /// </summary>
        void AddNewContent(NewContentType contentType, string block);
        /// <summary>
        /// Allows gouping content blocks by an adhoc identifier.
        /// </summary>
        void AddNewContent(string contentIdentifier, string block);
        /// <summary>
        /// Retrieves a block of content aggregated by <c>NewContentType</c>.
        /// </summary>
        /// <param name="newContentTypes"><c>NewContentType</c> blocks to aggregate</param>
        string RetrieveBlock(params NewContentType[] newContentTypes);
        /// <summary>
        /// Retrieves a block of content aggregated by a user-determined identifier.
        /// </summary>
        /// <param name="contentIdentifiers"><c>NewContentType</c> blocks to aggregate</param>
        string RetrieveBlock(params string[] contentIdentifiers);
        /// <summary>
        /// Sets default number of NewLines between blocks of code after
        /// all retrieving block(s) of code.  The default value is 2.
        /// </summary>
        int NewLineLimit { set; get; }
    }

    /// <summary>
    /// NewContentAggregator provides a repository for caching generated code blocks
    /// and retrieving them as an aggregated single block of code organized by <c>NewContentType</c>.
    /// </summary>
    public class NewContentAggregator : INewContentAggregator
    {
        private readonly Dictionary<NewContentType, List<string>> _newContent;
        private Dictionary<string, List<string>> _unStructuredContent;

        public NewContentAggregator()
        {
            _newContent = new Dictionary<NewContentType, List<string>>
            {
                { NewContentType.UserDefinedTypeDeclaration, new List<string>() },
                { NewContentType.DeclarationBlock, new List<string>() },
                { NewContentType.CodeSectionBlock, new List<string>() },
            };

            _unStructuredContent = new Dictionary<string, List<string>>();
        }

        public void AddNewContent(NewContentType contentType, string newContent)
        {
            if (!string.IsNullOrEmpty(newContent))
            {
                _newContent[contentType].Add(newContent);
            }
        }

        public void AddNewContent(string contentIdentifier, string newContent)
        {
            if (!string.IsNullOrEmpty(newContent))
            {
                if (!_unStructuredContent.ContainsKey(contentIdentifier))
                {
                    _unStructuredContent.Add(contentIdentifier, new List<string>());
                }
                _unStructuredContent[contentIdentifier].Add(newContent);
            }
        }

        public string RetrieveBlock(params NewContentType[] newContentTypes)
        {
            var block = string.Empty;
            foreach (var newContentType in newContentTypes)
            {
                var newContent = string.Join(NewLines.DOUBLE_SPACE, _newContent[newContentType]);
                if (!string.IsNullOrEmpty(newContent))
                {
                    block = string.IsNullOrEmpty(block)
                        ? newContent
                        : $"{block}{NewLines.DOUBLE_SPACE}{newContent}";
                }
            }
            return LimitNewLines(block.Trim(), NewLineLimit);
        }

        public string RetrieveBlock(params string[] contentIdentifiers)
        {
            var block = string.Empty;
            foreach (var identifier in contentIdentifiers)
            {
                if (_unStructuredContent.TryGetValue(identifier, out var adHocContent))
                {
                    var newContent = string.Join(NewLines.DOUBLE_SPACE, adHocContent);
                    if (!string.IsNullOrEmpty(newContent))
                    {
                        block = string.IsNullOrEmpty(block)
                            ? newContent
                            : $"{block}{NewLines.DOUBLE_SPACE}{newContent}";
                    }
                }
            }

            return string.IsNullOrEmpty(block)
                ? null
                : LimitNewLines(block.Trim(), NewLineLimit);
        }

        public int NewLineLimit { set; get; } = 2;

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
