using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    public enum ContentTypes { TypeDeclarationBlock, DeclarationBlock, MethodBlock, PostContentMessage };

    public interface IProvideNewContent
    {
        void AddMethod(string content);
        void AddFieldOrConstantDeclaration(string content);
        void AddTypeDeclaration(string content);
        void AddPostScriptComment(string content);
        string AsSingleBlock { get; }
        int CountOfLines { get; }
    }

    public class ContentToMove : IProvideNewContent
    {
        private Dictionary<ContentTypes, List<string>> _movedContent;

        public ContentToMove()
        {
            _movedContent = new Dictionary<ContentTypes, List<string>>
            {
                { ContentTypes.PostContentMessage, new List<string>() },
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
        public void AddPostScriptComment(string content) => Add(ContentTypes.PostContentMessage, content);

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
                                .Concat(_movedContent[ContentTypes.MethodBlock])
                                .Concat(_movedContent[ContentTypes.PostContentMessage]))
                                .Trim();
            }
        }

        public int CountOfLines 
            => AsSingleBlock.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).Count();
    }
}
