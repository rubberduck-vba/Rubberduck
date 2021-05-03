using Rubberduck.Refactorings.EncapsulateFieldUseBackingField;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    /// <summary>
    /// The EncapsulateFieldModel provides a facade to the EncapsulateFieldRefactoring
    /// by aggregating and simplifying access to the EncapsulateFieldUseBackingFieldModel 
    /// and the EncapsulateFieldUseBackingUDTMemberModel.
    /// </summary>
    public class EncapsulateFieldModel : IRefactoringModel
    {
        public EncapsulateFieldModel(EncapsulateFieldUseBackingFieldModel backingFieldModel,
            EncapsulateFieldUseBackingUDTMemberModel udtModel,
            IEncapsulateFieldConflictFinder conflictFinder) 
        {
            EncapsulateFieldUseBackingFieldModel = backingFieldModel;
            EncapsulateFieldUseBackingUDTMemberModel = udtModel;
            ObjectStateUDTCandidates = udtModel.ObjectStateUDTCandidates;
            ConflictFinder = conflictFinder;
            EncapsulateFieldUseBackingFieldModel.ConflictFinder = conflictFinder;
            EncapsulateFieldUseBackingUDTMemberModel.ConflictFinder = conflictFinder;
        }

        public EncapsulateFieldUseBackingUDTMemberModel EncapsulateFieldUseBackingUDTMemberModel { get; }

        public EncapsulateFieldUseBackingFieldModel EncapsulateFieldUseBackingFieldModel { get; }

        public string PreviewRefactoring() => PreviewProvider?.Preview(this) ?? string.Empty;

        public IRefactoringPreviewProvider<EncapsulateFieldModel> PreviewProvider { set; get; }

        public Action<EncapsulateFieldModel> StrategyChangedAction { set; get; } = (m) => { };

        public Action<EncapsulateFieldModel> ObjectStateFieldChangedAction { set; get; } = (m) => { };

        public IReadOnlyCollection<IObjectStateUDT> ObjectStateUDTCandidates { private set; get; }

        public IEncapsulateFieldConflictFinder ConflictFinder { set; get; }

        public IObjectStateUDT ObjectStateUDTField
        {
            set
            {
                if (EncapsulateFieldUseBackingUDTMemberModel.ObjectStateUDTField != value)
                {
                    EncapsulateFieldUseBackingUDTMemberModel.ObjectStateUDTField = value;
                    ObjectStateFieldChangedAction(this);
                }
            }
            get => EncapsulateFieldStrategy == EncapsulateFieldStrategy.ConvertFieldsToUDTMembers
                ? EncapsulateFieldUseBackingUDTMemberModel.ObjectStateUDTField
                : null;
        }

        private EncapsulateFieldStrategy _strategy;
        public EncapsulateFieldStrategy EncapsulateFieldStrategy
        {
            set
            {
                if (_strategy != value)
                {
                    _strategy = value;
                    StrategyChangedAction(this);
                }
            }
            get => _strategy;
        }

        public IReadOnlyCollection<IEncapsulateFieldCandidate> EncapsulationCandidates => EncapsulateFieldStrategy == EncapsulateFieldStrategy.UseBackingFields
            ? EncapsulateFieldUseBackingFieldModel.EncapsulationCandidates
            : EncapsulateFieldUseBackingUDTMemberModel.EncapsulationCandidates;

        public IEnumerable<IEncapsulateFieldCandidate> SelectedFieldCandidates
            => EncapsulationCandidates.Where(v => v.EncapsulateFlag);

        public IEncapsulateFieldCandidate this[string encapsulatedFieldTargetID]
            => EncapsulationCandidates.Where(c => c.TargetID.Equals(encapsulatedFieldTargetID)).Single();
    }
}
