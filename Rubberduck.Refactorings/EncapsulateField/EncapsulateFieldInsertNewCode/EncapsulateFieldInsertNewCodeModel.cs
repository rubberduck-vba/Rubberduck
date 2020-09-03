using Rubberduck.VBEditor;
using Rubberduck.Refactorings.CodeBlockInsert;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Refactorings.EncapsulateField;

namespace Rubberduck.Refactorings.EncapsulateFieldInsertNewCode
{
    public class EncapsulateFieldInsertNewCodeModel : IRefactoringModel
    {
        public EncapsulateFieldInsertNewCodeModel(IEnumerable<IEncapsulateFieldCandidate> selectedFieldCandidates)
        {
            _selectedCandidates = selectedFieldCandidates.ToList();
            if (_selectedCandidates.Any())
            {
                QualifiedModuleName = _selectedCandidates.Select(f => f.QualifiedModuleName).First();
            }

            NewContent = new Dictionary<NewContentType, List<string>>
            {
                { NewContentType.PostContentMessage, new List<string>() },
                { NewContentType.DeclarationBlock, new List<string>() },
                { NewContentType.CodeSectionBlock, new List<string>() },
                { NewContentType.TypeDeclarationBlock, new List<string>() }
            };
        }

        public bool IncludeNewContentMarker { set; get; } = false;

        public QualifiedModuleName QualifiedModuleName { get; } = new QualifiedModuleName();

        public Dictionary<NewContentType, List<string>> NewContent { set; get; }

        private List<IEncapsulateFieldCandidate> _selectedCandidates;
        public IEnumerable<IEncapsulateFieldCandidate> SelectedFieldCandidates
            => _selectedCandidates;
    }
}
