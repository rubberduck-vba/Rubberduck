using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public enum NewContentType
    {
        ImplementsDeclaration,
        WithEventsDeclaration,
        EnumerationTypeDeclaration,
        UserDefinedTypeDeclaration,
        DeclarationBlock,
        CodeSectionBlock,
    }

    public interface INewContentAggregator
    {
        void AddNewContent(NewContentType contentType, string block);
        void AddNewContent(string contentIdentifier, string block);
        string RetrieveBlock(params NewContentType[] newContentTypes);
        string RetrieveBlock(params string[] contentIdentifiers);
        int NewLineLimit { set; get; }
    }

    public class NewContentAggregator : INewContentAggregator
    {
        private readonly static string _doubleSpace = $"{Environment.NewLine}{Environment.NewLine}";
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
                var newContent = string.Join(_doubleSpace, _newContent[newContentType]);
                if (!string.IsNullOrEmpty(newContent))
                {
                    block = string.IsNullOrEmpty(block)
                        ? newContent
                        : $"{block}{_doubleSpace}{newContent}";
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
                    var newContent = string.Join(_doubleSpace, adHocContent);
                    if (!string.IsNullOrEmpty(newContent))
                    {
                        block = string.IsNullOrEmpty(block)
                            ? newContent
                            : $"{block}{_doubleSpace}{newContent}";
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
