using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    public enum ContentTypes { TypeDeclarationBlock, DeclarationBlock, MethodBlock };

    public interface INewContentProvider
    {
        void AddMethod(string content);
        void AddFieldOrConstantDeclaration(string content);
        void AddTypeDeclaration(string content);
        string AsSingleBlock { get; }
        string AsSingleBlockWithinDemarcationComments(string startMessage = null, string endMessage = null);
    }

    public class ContentToMove : INewContentProvider
    {
        private Dictionary<ContentTypes, List<string>> _movedContent;

        public ContentToMove()
        {
            _movedContent = new Dictionary<ContentTypes, List<string>>
            {
                { ContentTypes.DeclarationBlock, new List<string>() },
                { ContentTypes.MethodBlock, new List<string>() },
                { ContentTypes.TypeDeclarationBlock, new List<string>() }
            };
        }

        public string this[ContentTypes contentType]
        {
            get
            {
                return string.Empty;
            }
        }

        public void AddMethod(string content) => Add(ContentTypes.MethodBlock, content);
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
                return string.Join($"{ Environment.NewLine}{ Environment.NewLine}",
                                (_movedContent[ContentTypes.TypeDeclarationBlock])
                                .Concat(_movedContent[ContentTypes.DeclarationBlock])
                                .Concat(_movedContent[ContentTypes.MethodBlock]))
                                .Trim();
            }
        }

        public string AsSingleBlockWithinDemarcationComments(string startMessage = null, string endMessage = null)
        {
            var changesStartMarker = startMessage ?? MoveMemberResources.MovedContentBelowThisLine;
            var changesEndMarker = endMessage ?? MoveMemberResources.MovedContentAboveThisLine;

            return $"'***** {changesStartMarker}  *****{Environment.NewLine}{AsSingleBlock}{Environment.NewLine}'**** {changesEndMarker} ****{Environment.NewLine}";
        }
    }
}
