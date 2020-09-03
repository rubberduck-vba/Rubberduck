using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingField;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldModel : IRefactoringModel
    {
        public EncapsulateFieldModel(EncapsulateFieldUseBackingFieldModel backingFieldModel,
            EncapsulateFieldUseBackingUDTMemberModel udtModel)
        {
            EncapsulateFieldUseBackingFieldModel = backingFieldModel;
            EncapsulateFieldUseBackingUDTMemberModel = udtModel;
            ResetConflictDetection(EncapsulateFieldStrategy.UseBackingFields);
        }

        public EncapsulateFieldUseBackingUDTMemberModel EncapsulateFieldUseBackingUDTMemberModel { get; }

        public EncapsulateFieldUseBackingFieldModel EncapsulateFieldUseBackingFieldModel { get; }

        public IRefactoringPreviewProvider<EncapsulateFieldModel> PreviewProvider { set; get; }

        public IEnumerable<IObjectStateUDT> ObjectStateUDTCandidates => EncapsulateFieldUseBackingUDTMemberModel.ObjectStateUDTCandidates;

        private EncapsulateFieldStrategy _strategy;
        public EncapsulateFieldStrategy EncapsulateFieldStrategy
        {
            set
            {
                if (_strategy == value)
                {
                    return;
                }
                _strategy = value;
                ResetConflictDetection(_strategy);
            }
            get => _strategy;
        }

        private void ResetConflictDetection(EncapsulateFieldStrategy strategy)
        {
            var conflictFinder =
                strategy == EncapsulateFieldStrategy.UseBackingFields
                    ? EncapsulateFieldUseBackingFieldModel.ConflictFinder
                    : EncapsulateFieldUseBackingUDTMemberModel.ConflictFinder;

            foreach (var candidate in EncapsulateFieldUseBackingFieldModel.EncapsulationCandidates)
            {
                candidate.ConflictFinder = conflictFinder;
                ResolveConflict(conflictFinder, candidate);
            }

            return;
        }

        private void ResolveConflict(IEncapsulateFieldConflictFinder conflictFinder, IEncapsulateFieldCandidate candidate)
        {
            conflictFinder.AssignNoConflictIdentifiers(candidate);
            if (candidate is IUserDefinedTypeCandidate udtCandidate)
            {
                foreach (var member in udtCandidate.Members)
                {
                    conflictFinder.AssignNoConflictIdentifiers(member);
                    if (member.WrappedCandidate is IUserDefinedTypeCandidate childUDT
                        && childUDT.Declaration.AsTypeDeclaration.HasPrivateAccessibility())
                    {
                        ResolveConflict(conflictFinder, childUDT);
                    }
                }
            }
        }

        public IReadOnlyCollection<IEncapsulateFieldCandidate> EncapsulationCandidates => EncapsulateFieldStrategy == EncapsulateFieldStrategy.UseBackingFields
            ? EncapsulateFieldUseBackingFieldModel.EncapsulationCandidates
            : EncapsulateFieldUseBackingUDTMemberModel.EncapsulationCandidates;

        public IEnumerable<IEncapsulateFieldCandidate> SelectedFieldCandidates
            => EncapsulationCandidates.Where(v => v.EncapsulateFlag);

        public IEncapsulateFieldCandidate this[string encapsulatedFieldTargetID]
            => EncapsulationCandidates.Where(c => c.TargetID.Equals(encapsulatedFieldTargetID)).Single();

        public IObjectStateUDT ObjectStateUDTField
        {
            get => EncapsulateFieldUseBackingUDTMemberModel.ObjectStateUDTField;
            set => EncapsulateFieldUseBackingUDTMemberModel.ObjectStateUDTField = value;
        }
    }
}
