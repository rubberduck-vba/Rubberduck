using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldModel : IRefactoringModel
    {
        private readonly Func<EncapsulateFieldModel, string> _previewDelegate;
        private QualifiedModuleName _targetQMN;
        private IDeclarationFinderProvider _declarationFinderProvider;
        private IEncapsulateFieldValidationsProvider _validationsProvider;
        private IObjectStateUDT _newObjectStateUDT;

        private List<IEncapsulateFieldCandidate> _convertedFields;
        private HashSet<IObjectStateUDT> _objStateCandidates;

        private IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> _udtFieldToUdtDeclarationMap = new Dictionary<Declaration, (Declaration, IEnumerable<Declaration>)>();

        public EncapsulateFieldModel(
            Declaration target, 
            IEnumerable<IEncapsulateFieldCandidate> candidates, 
            IEnumerable<IObjectStateUDT> objectStateUDTCandidates, 
            IObjectStateUDT stateUDTField, 
            Func<EncapsulateFieldModel, string> previewDelegate, 
            IDeclarationFinderProvider declarationFinderProvider, 
            IEncapsulateFieldValidationsProvider validationsProvider)
        {
            _previewDelegate = previewDelegate;
            _targetQMN = target.QualifiedModuleName;
            _newObjectStateUDT = stateUDTField;
            _declarationFinderProvider = declarationFinderProvider;
            _validationsProvider = validationsProvider;

            _useBackingFieldCandidates = candidates.ToList();
            _objStateCandidates = new HashSet<IObjectStateUDT>(objectStateUDTCandidates);
            _objStateCandidates.Add(_newObjectStateUDT);

            EncapsulateFieldStrategy = EncapsulateFieldStrategy.UseBackingFields;
            _activeObjectStateUDT = ObjectStateUDTField;
        }

        public QualifiedModuleName QualifiedModuleName => _targetQMN;

        public string PreviewRefactoring() => _previewDelegate(this);

        public IEnumerable<IObjectStateUDT> ObjectStateUDTCandidates => _objStateCandidates;

        private EncapsulateFieldStrategy _encapsulationFieldStategy;
        public EncapsulateFieldStrategy EncapsulateFieldStrategy
        {
            get => _encapsulationFieldStategy;
            set
            {
                if (_encapsulationFieldStategy == value) { return; }

                _encapsulationFieldStategy = value;

                if (_encapsulationFieldStategy == EncapsulateFieldStrategy.UseBackingFields)
                {
                    UpdateFieldCandidatesForUseBackingFieldsStrategy();
                    return;
                }
                UpdateFieldCandidatesForConvertFieldsToUDTMembersStrategy();
            }
        } 

        public IEncapsulateFieldValidationsProvider ValidationsProvider => _validationsProvider;

        private List<IEncapsulateFieldCandidate> _useBackingFieldCandidates;
        public List<IEncapsulateFieldCandidate> EncapsulationCandidates
        {
            get
            {
                if (EncapsulateFieldStrategy == EncapsulateFieldStrategy.UseBackingFields)
                {
                    return _useBackingFieldCandidates;
                }

                if (_convertedFields is null)
                {
                    _convertedFields = new List<IEncapsulateFieldCandidate>();
                    foreach (var field in _useBackingFieldCandidates)
                    {
                        _convertedFields.Add(new ConvertToUDTMember(field, ObjectStateUDTField));
                    }
                }
                return _convertedFields;
            }
        } 

        public IEnumerable<IEncapsulateFieldCandidate> SelectedFieldCandidates
            => EncapsulationCandidates.Where(v => v.EncapsulateFlag);

        public IEnumerable<IUserDefinedTypeCandidate> UDTFieldCandidates 
            => EncapsulationCandidates
                    .Where(v => v is IUserDefinedTypeCandidate)
                    .Cast<IUserDefinedTypeCandidate>();

        public IEnumerable<IUserDefinedTypeCandidate> SelectedUDTFieldCandidates 
            => SelectedFieldCandidates
                    .Where(v => v is IUserDefinedTypeCandidate)
                    .Cast<IUserDefinedTypeCandidate>();

        public IEncapsulateFieldCandidate this[string encapsulatedFieldTargetID]
            => EncapsulationCandidates.Where(c => c.TargetID.Equals(encapsulatedFieldTargetID)).Single();

        public IEncapsulateFieldCandidate this[Declaration fieldDeclaration]
            => EncapsulationCandidates.Where(c => c.Declaration == fieldDeclaration).Single();
        
        private IObjectStateUDT _activeObjectStateUDT;
        public IObjectStateUDT ObjectStateUDTField
        {
            get
            {
                _activeObjectStateUDT = ObjectStateUDTCandidates
                            .SingleOrDefault(os => os.IsSelected) ?? _newObjectStateUDT;

                return _activeObjectStateUDT;
            }
            set
            {
                if (_activeObjectStateUDT.FieldIdentifier == (value?.FieldIdentifier ?? string.Empty))
                {
                    return;
                }

                foreach (var objectStateUDT in ObjectStateUDTCandidates)
                {
                    objectStateUDT.IsSelected = false;
                }

                _activeObjectStateUDT =
                    ObjectStateUDTCandidates.SingleOrDefault(os => os.FieldIdentifier.Equals(value?.FieldIdentifier ?? null))
                    ?? _newObjectStateUDT;

                _activeObjectStateUDT.IsSelected = true;

                if (EncapsulateFieldStrategy == EncapsulateFieldStrategy.ConvertFieldsToUDTMembers)
                {
                    foreach (var field in EncapsulationCandidates)
                    {
                        if (field is IConvertToUDTMember convertedField)
                        {
                            convertedField.ObjectStateUDT = _activeObjectStateUDT;
                            convertedField.ConflictFinder = _validationsProvider.ConflictDetector(EncapsulateFieldStrategy, _declarationFinderProvider);
                            convertedField.ConflictFinder.AssignNoConflictIdentifiers(convertedField);
                        }
                    }
                }
            }
        }

        private void UpdateFieldCandidatesForUseBackingFieldsStrategy()
        {
            foreach (var candidate in EncapsulationCandidates)
            {
                switch (candidate)
                {
                    case IUserDefinedTypeCandidate udt:
                        candidate.NameValidator = EncapsulateFieldValidationsProvider.NameOnlyValidator(NameValidators.UserDefinedType);
                        break;
                    case IUserDefinedTypeMemberCandidate udtm:
                        candidate.NameValidator = candidate.Declaration.IsArray
                            ? EncapsulateFieldValidationsProvider.NameOnlyValidator(NameValidators.UserDefinedTypeMemberArray)
                            : EncapsulateFieldValidationsProvider.NameOnlyValidator(NameValidators.UserDefinedTypeMember);
                        break;
                    default:
                        candidate.NameValidator = EncapsulateFieldValidationsProvider.NameOnlyValidator(NameValidators.Default);
                        break;
                }
                candidate.ConflictFinder = _validationsProvider.ConflictDetector(EncapsulateFieldStrategy, _declarationFinderProvider);
                candidate.ConflictFinder.AssignNoConflictIdentifiers(candidate);
            }
        }

        private void UpdateFieldCandidatesForConvertFieldsToUDTMembersStrategy()
        {
            foreach (var candidate in EncapsulationCandidates.Cast<IConvertToUDTMember>())
            {
                candidate.ObjectStateUDT = ObjectStateUDTField;
                candidate.NameValidator = candidate.Declaration.IsArray
                    ? EncapsulateFieldValidationsProvider.NameOnlyValidator(NameValidators.UserDefinedTypeMemberArray)
                    : EncapsulateFieldValidationsProvider.NameOnlyValidator(NameValidators.UserDefinedTypeMember);

                candidate.ConflictFinder = _validationsProvider.ConflictDetector(EncapsulateFieldStrategy, _declarationFinderProvider);
                candidate.ConflictFinder.AssignNoConflictIdentifiers(candidate);
            }
        }
    }
}
