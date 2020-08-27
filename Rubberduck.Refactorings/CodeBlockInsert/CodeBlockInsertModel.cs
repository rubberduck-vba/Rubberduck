using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.CodeBlockInsert
{
    public enum NewContentType
    {
        TypeDeclarationBlock,
        DeclarationBlock,
        CodeSectionBlock,
        PostContentMessage
    }

    public class CodeBlockInsertModel : IRefactoringModel
    {
        public CodeBlockInsertModel()
        {
            _selectedFields = new List<IEncapsulateFieldCandidate>();

            NewContent = new Dictionary<NewContentType, List<string>>
            {
                { NewContentType.PostContentMessage, new List<string>() },
                { NewContentType.DeclarationBlock, new List<string>() },
                { NewContentType.CodeSectionBlock, new List<string>() },
                { NewContentType.TypeDeclarationBlock, new List<string>() }
            };
        }

        public QualifiedModuleName QualifiedModuleName { set; get; }

        public bool IncludeComments { set; get; }

        public Dictionary<NewContentType, List<string>> NewContent { set; get; }

        public void AddContentBlock(NewContentType contentType, string block)
            => NewContent[contentType].Add(block);

        public int NewLineLimit { set; get; } = 2;

        private List<IEncapsulateFieldCandidate> _selectedFields;
        public IEnumerable<IEncapsulateFieldCandidate> SelectedFieldCandidates
        {
            set => _selectedFields = value.ToList();
            get => _selectedFields;
        }

        public int? CodeSectionStartIndex { set; get; } = null;

        public int? NewContentInsertionIndex => CodeSectionStartIndex;
    }
}
