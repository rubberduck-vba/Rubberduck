using Rubberduck.VBEditor;
using Rubberduck.Refactorings.CodeBlockInsert;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
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
        }

        public bool IncludeNewContentMarker { set; get; } = false;

        public QualifiedModuleName QualifiedModuleName { get; } = new QualifiedModuleName();

        private Dictionary<NewContentType, List<string>> _newContent { set; get; }
        public Dictionary<NewContentType, List<string>> NewContent
        {
            set => _newContent = value;
            get => _newContent;
        }

        private List<IEncapsulateFieldCandidate> _selectedCandidates;
        public IEnumerable<IEncapsulateFieldCandidate> SelectedFieldCandidates
            => _selectedCandidates;
        //{
        //    //set => _selectedCandidates = value.ToList();
        //    get => _selectedCandidates;
        //}
    }
}
