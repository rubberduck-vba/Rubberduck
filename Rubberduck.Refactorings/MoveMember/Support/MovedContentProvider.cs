using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember
{
    public enum ContentTypes
    {
        TypeDeclarationBlock,
        DeclarationBlock,
        MethodBlock
    };

    public interface IMovedContentProvider
    {
        void AddMethodDeclaration(string content);
        void AddFieldOrConstantDeclaration(string content);
        void AddTypeDeclaration(string content);
        string AsSingleBlock { get; }
    }

    public interface IMovedContentPreviewProvider : IMovedContentProvider
    { }

    public class MovedContentProvider : MovedContentProviderBase, IMovedContentProvider
    {
        public MovedContentProvider()
               : base(false) { }
    }

    public class MovedContentPreviewProvider : MovedContentProviderBase, IMovedContentPreviewProvider
    {
        public MovedContentPreviewProvider()
            : base(true) { }
    }

    public class MovedContentProviderBase
    {
        private Dictionary<ContentTypes, List<string>> _movedContent;
        private readonly bool _applyPreviewAnnotations;

        public MovedContentProviderBase(bool applyPreviewAnnotations)
        {
            _applyPreviewAnnotations = applyPreviewAnnotations;
            _movedContent = new Dictionary<ContentTypes, List<string>>
            {
                { ContentTypes.DeclarationBlock, new List<string>() },
                { ContentTypes.MethodBlock, new List<string>() },
                { ContentTypes.TypeDeclarationBlock, new List<string>() }
            };
        }

        public void AddMethodDeclaration(string content) => Add(ContentTypes.MethodBlock, content);
        public void AddFieldOrConstantDeclaration(string content) => Add(ContentTypes.DeclarationBlock, content);
        public void AddTypeDeclaration(string content) => Add(ContentTypes.TypeDeclarationBlock, content);

        private void Add(ContentTypes contentType, string content)
        {
            if (_movedContent.TryGetValue(contentType, out var blocks))
            {
                blocks.Add(content);
                return;
            }
            _movedContent.Add(contentType, new List<string>() { content });
        }

        public string AsSingleBlock
        {
            get
            {
                var result = string.Join($"{ Environment.NewLine}{ Environment.NewLine}",
                                (_movedContent[ContentTypes.TypeDeclarationBlock])
                                .Concat(_movedContent[ContentTypes.DeclarationBlock])
                                .Concat(_movedContent[ContentTypes.MethodBlock]))
                                .Trim();

                if (_applyPreviewAnnotations)
                {
                    return $"'*****  {Resources.RubberduckUI.MoveMember_MovedContentBelowThisLine}  *****{Environment.NewLine}{Environment.NewLine}{result}{Environment.NewLine}{Environment.NewLine}'****  {Resources.RubberduckUI.MoveMember_MovedContentAboveThisLine}  ****{Environment.NewLine}";
                }
                return result;
            }
        }
    }
}
