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
        bool HasContent { get; }
    }

    public interface INewContentPreviewProvider : INewContentProvider
    {
        string AddNewContentDemarcation(string block);
    }

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

        private Dictionary<NewContentBlocks, List<string>> _newContent;
        private readonly bool _applyPreviewAnnotations;

        private static string _blockSpacing = $"{Environment.NewLine}{Environment.NewLine}";
        private static string _markers = "*****";

        private static string _annotationStartMsg = Resources.RubberduckUI.MoveMember_NewContentBelowThisLine;
        private static string _annotationsEndMsg = Resources.RubberduckUI.MoveMember_NewContentAboveThisLine;

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
            _newContent = new Dictionary<NewContentBlocks, List<string>>
            {
                { NewContentBlocks.FieldOrConstantDeclaration, new List<string>() },
                { NewContentBlocks.Member, new List<string>() },
                { NewContentBlocks.TypeDeclaration, new List<string>() }
            };
            return this as INewContentProvider;
        }

        private void Add(NewContentBlocks contentType, string content)
        {
            if (_newContent.TryGetValue(contentType, out var blocks))
            {
                blocks.Add(content);
                return;
            }
            _newContent.Add(contentType, new List<string>() { content });
        }

        public bool HasContent => _newContent.Values.Any(v => v.Any());

        public string AsSingleBlock
        {
            get
            {
                if (!HasContent)
                {
                    return string.Empty;
                }

                var result = string.Join(_blockSpacing,
                                (_newContent[NewContentBlocks.TypeDeclaration])
                                .Concat(_newContent[NewContentBlocks.FieldOrConstantDeclaration])
                                .Concat(_newContent[NewContentBlocks.Member]))
                                .Trim();

                return _applyPreviewAnnotations
                            ? AddNewContentDemarcation(result)
                            : result;
            }
        }

        public string AddNewContentDemarcation(string block)
        {
            var contentLines = new List<string>()
                {
                    $"'{_markers}  {_annotationStartMsg}  {_markers}",
                    block,
                    $"'{_markers}  {_annotationsEndMsg}  {_markers}"
                };

            return string.Join(_blockSpacing, contentLines);
        }
    }
}
