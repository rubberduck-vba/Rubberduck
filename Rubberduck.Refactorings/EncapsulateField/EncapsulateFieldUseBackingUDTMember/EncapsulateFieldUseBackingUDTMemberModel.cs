using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.Refactorings.EncapsulateField;

namespace Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember
{
    public class EncapsulateFieldUseBackingUDTMemberModel : IRefactoringModel
    {
        private readonly List<IEncapsulateFieldAsUDTMemberCandidate> _encapsulateAsUDTMemberCandidates;

        public EncapsulateFieldUseBackingUDTMemberModel(IObjectStateUDT targetObjectStateUserDefinedTypeField, 
            IEnumerable<IEncapsulateFieldAsUDTMemberCandidate> encapsulateAsUDTMemberCandidates,
            IEnumerable<IObjectStateUDT> objectStateUserDefinedTypeCandidates)
        {
            _encapsulateAsUDTMemberCandidates = encapsulateAsUDTMemberCandidates.ToList();

            EncapsulationCandidates = _encapsulateAsUDTMemberCandidates.Cast<IEncapsulateFieldCandidate>().ToList();

            ObjectStateUDTField = targetObjectStateUserDefinedTypeField;

            ObjectStateUDTCandidates = objectStateUserDefinedTypeCandidates.ToList();

            QualifiedModuleName = encapsulateAsUDTMemberCandidates.First().QualifiedModuleName;

        }

        public INewContentAggregator NewContentAggregator { set; get; }

        public IReadOnlyCollection<IObjectStateUDT> ObjectStateUDTCandidates { get; }

        public IEncapsulateFieldConflictFinder ConflictFinder { set; get; }

        public IReadOnlyCollection<IEncapsulateFieldCandidate> EncapsulationCandidates { get; }

        public IReadOnlyCollection<IEncapsulateFieldCandidate> SelectedFieldCandidates
            => _encapsulateAsUDTMemberCandidates
                .Where(v => v.EncapsulateFlag)
                .ToList();

        public QualifiedModuleName QualifiedModuleName { get; }

        public IObjectStateUDT ObjectStateUDTField
        {
            set
            {
                if (ObjectStateUDTField != null)
                {
                    ObjectStateUDTField.IsSelected = false;
                }

                if (value != null)
                {
                    value.IsSelected = true;
                }

                _encapsulateAsUDTMemberCandidates.ForEach(cf => cf.ObjectStateUDT = value);
            }
            get => _encapsulateAsUDTMemberCandidates.FirstOrDefault()?.ObjectStateUDT;
        }
    }
}
