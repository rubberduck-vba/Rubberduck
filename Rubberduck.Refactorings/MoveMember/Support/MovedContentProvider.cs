using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface INewContentProvider
    {
        void AddMember(string content);
        void AddFieldOrConstantDeclaration(string content);
        void AddTypeDeclaration(string content);
        INewContentProvider ResetContent();
        string AsSingleBlock { get; }
    }

    public interface INewContentPreviewProvider : INewContentProvider
    { }

    public class NewContentProvider : NewContentProviderBase, INewContentProvider
    {
        public NewContentProvider()
               : base(false) { }
    }

    public class NewContentPreviewProvider : NewContentProviderBase, INewContentPreviewProvider
    {
        public NewContentPreviewProvider()
            : base(true) { }
    }

    public class NewContentProviderBase
    {
        private enum NewContentBlocks
        {
            TypeDeclaration,
            FieldOrConstantDeclaration,
            Member
        };

        private Dictionary<NewContentBlocks, List<string>> _movedContent;
        private readonly bool _applyPreviewAnnotations;

        public NewContentProviderBase(bool applyPreviewAnnotations)
        {
            _applyPreviewAnnotations = applyPreviewAnnotations;
            ResetContent();
        }

        public void AddMember(string content) => Add(NewContentBlocks.Member, content);

        public void AddFieldOrConstantDeclaration(string content) => Add(NewContentBlocks.FieldOrConstantDeclaration, content);

        public void AddTypeDeclaration(string content) => Add(NewContentBlocks.TypeDeclaration, content);

        public INewContentProvider ResetContent()
        {
            _movedContent = new Dictionary<NewContentBlocks, List<string>>
            {
                { NewContentBlocks.FieldOrConstantDeclaration, new List<string>() },
                { NewContentBlocks.Member, new List<string>() },
                { NewContentBlocks.TypeDeclaration, new List<string>() }
            };
            return this as INewContentProvider;
        }

        private void Add(NewContentBlocks contentType, string content)
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
                                (_movedContent[NewContentBlocks.TypeDeclaration])
                                .Concat(_movedContent[NewContentBlocks.FieldOrConstantDeclaration])
                                .Concat(_movedContent[NewContentBlocks.Member]))
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
