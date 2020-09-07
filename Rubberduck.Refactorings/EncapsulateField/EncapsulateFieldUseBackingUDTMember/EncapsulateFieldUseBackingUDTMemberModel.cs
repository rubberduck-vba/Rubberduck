using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.Refactorings.CodeBlockInsert;
using Rubberduck.Refactorings.EncapsulateField;

namespace Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember
{
    public class EncapsulateFieldUseBackingUDTMemberModel : IRefactoringModel
    {
        private List<IEncapsulateFieldAsUDTMemberCandidate> _encapsulateAsUDTMemberCandidates;

        public EncapsulateFieldUseBackingUDTMemberModel(IObjectStateUDT targetObjectStateUserDefinedTypeField, 
            IEnumerable<IEncapsulateFieldAsUDTMemberCandidate> encapsulateAsUDTMemberCandidates,
            IEnumerable<IObjectStateUDT> objectStateUserDefinedTypeCandidates)
        {
            _encapsulateAsUDTMemberCandidates = new List<IEncapsulateFieldAsUDTMemberCandidate>(encapsulateAsUDTMemberCandidates);

            ObjectStateUDTField = targetObjectStateUserDefinedTypeField;

            ObjectStateUDTCandidates = objectStateUserDefinedTypeCandidates.ToList();

            QualifiedModuleName = encapsulateAsUDTMemberCandidates.First().QualifiedModuleName;

            ResetNewContent();
        }

        public void ResetNewContent()
        {
            NewContent = new Dictionary<NewContentType, List<string>>
            {
                { NewContentType.PostContentMessage, new List<string>() },
                { NewContentType.DeclarationBlock, new List<string>() },
                { NewContentType.CodeSectionBlock, new List<string>() },
                { NewContentType.TypeDeclarationBlock, new List<string>() }
            };
        }

        public IReadOnlyCollection<IObjectStateUDT> ObjectStateUDTCandidates { get; }

        public IEncapsulateFieldConflictFinder ConflictFinder { set; get; }

        public bool IncludeNewContentMarker { set; get; } = false;

        public IReadOnlyCollection<IEncapsulateFieldCandidate> EncapsulationCandidates
            => _encapsulateAsUDTMemberCandidates.Cast<IEncapsulateFieldCandidate>().ToList();

        public IEnumerable<IEncapsulateFieldCandidate> SelectedFieldCandidates
            => EncapsulationCandidates.Where(v => v.EncapsulateFlag && v.Declaration != ObjectStateUDTField?.Declaration);

        public void AddContentBlock(NewContentType contentType, string block)
            => NewContent[contentType].Add(block);

        public Dictionary<NewContentType, List<string>> NewContent { set; get; }

        public QualifiedModuleName QualifiedModuleName { get; }

        private IObjectStateUDT _objectStateUDT;
        public IObjectStateUDT ObjectStateUDTField
        {
            set
            {
                if (_objectStateUDT == value)
                {
                    return;
                }

                if (_objectStateUDT != null)
                {
                    _objectStateUDT.IsSelected = false;
                }

                _objectStateUDT = value;
                if (_objectStateUDT != null)
                {
                    _objectStateUDT.IsSelected = true;
                }
                _encapsulateAsUDTMemberCandidates.ForEach(cf => cf.ObjectStateUDT = _objectStateUDT);
            }
            get => _objectStateUDT;
        }
    }
}
