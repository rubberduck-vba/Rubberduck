using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldModel : IRefactoringModel
    {
        private readonly Func<EncapsulateFieldModel, string> _previewDelegate;
        private QualifiedModuleName _targetQMN;
        //private IValidateEncapsulateFieldNames _validator;
        private IDeclarationFinderProvider _declarationFinderProvider;
        private IEncapsulateFieldValidationsProvider _validatorProvider;
        private IObjectStateUDT _newObjectStateUDT;

        private IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> _udtFieldToUdtDeclarationMap = new Dictionary<Declaration, (Declaration, IEnumerable<Declaration>)>();

        public EncapsulateFieldModel(Declaration target, IEnumerable<IEncapsulateFieldCandidate> candidates, IObjectStateUDT stateUDTField, Func<EncapsulateFieldModel, string> previewDelegate, IDeclarationFinderProvider declarationFinderProvider, IEncapsulateFieldValidationsProvider validatorProvider) // IEncapsulateFieldValidator validator)
        {
            _previewDelegate = previewDelegate;
            _targetQMN = target.QualifiedModuleName;
            _newObjectStateUDT = stateUDTField;
            _declarationFinderProvider = declarationFinderProvider;
            _validatorProvider = validatorProvider;

            EncapsulationCandidates = candidates.ToList();
            EncapsulateFieldStrategy = EncapsulateFieldStrategy.UseBackingFields;
        }

        public QualifiedModuleName QualifiedModuleName => _targetQMN;

        private EncapsulateFieldStrategy _encapsulationFieldStategy;
        public EncapsulateFieldStrategy EncapsulateFieldStrategy
        {
            set
            {
                _encapsulationFieldStategy = value;
                AssignCandidateValidations(value);
            }
            get => _encapsulationFieldStategy;
        } 

        public IEncapsulateFieldValidationsProvider ValidatorProvider => _validatorProvider;

        public IEncapsulateFieldConflictFinder ConflictDetector
            => _validatorProvider.ConflictDetector(EncapsulateFieldStrategy, _declarationFinderProvider);

        public List<IEncapsulateFieldCandidate> EncapsulationCandidates { set; get; } = new List<IEncapsulateFieldCandidate>();

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
        
        ////TODO: Get rid of this property
        //private bool _convertFieldsToUDTMembers;
        //public bool ConvertFieldsToUDTMembers
        //{
        //    set
        //    {
        //        _convertFieldsToUDTMembers = value;

        //        EncapsulateFieldStrategy = value
        //            ? EncapsulateFieldStrategy.ConvertFieldsToUDTMembers
        //            : EncapsulateFieldStrategy.UseBackingFields;
        //    }
        //    get => _convertFieldsToUDTMembers;
        //}

        private IObjectStateUDT _activeObjectStateUDT;
        public IObjectStateUDT StateUDTField
        {
            set
            {
                _activeObjectStateUDT = value;
                foreach (var candidate in EncapsulationCandidates)
                {
                    if (candidate is IConvertToUDTMember udtMember)
                    {
                        udtMember.ObjectStateUDT = value;
                    }
                }
            }
            get => _activeObjectStateUDT ?? _newObjectStateUDT;
        }

        public void AssignCandidateValidations(EncapsulateFieldStrategy strategy)
        {
            foreach (var candidate in EncapsulationCandidates)
            {
                candidate.ConvertFieldToUDTMember = strategy == EncapsulateFieldStrategy.ConvertFieldsToUDTMembers;

                candidate.ConflictFinder = ConflictDetector;
                if (strategy == EncapsulateFieldStrategy.UseBackingFields)
                {
                    if (candidate is IUserDefinedTypeCandidate)
                    {
                        candidate.NameValidator = _validatorProvider.NameOnlyValidator(Validators.UserDefinedType);
                    }
                    else if (candidate is IUserDefinedTypeMemberCandidate)
                    {
                        candidate.NameValidator = candidate.Declaration.IsArray
                            ? _validatorProvider.NameOnlyValidator(Validators.UserDefinedTypeMemberArray)
                            : _validatorProvider.NameOnlyValidator(Validators.UserDefinedTypeMember);
                    }
                    else
                    {
                        candidate.NameValidator = _validatorProvider.NameOnlyValidator(Validators.Default);
                    }
                }
                else
                {
                    candidate.NameValidator = candidate.Declaration.IsArray
                        ? _validatorProvider.NameOnlyValidator(Validators.UserDefinedTypeMemberArray)
                        : _validatorProvider.NameOnlyValidator(Validators.UserDefinedTypeMember);
                }
            }
        }

        public string PreviewRefactoring() => _previewDelegate(this);

        private HashSet<IObjectStateUDT>  _objStateCandidates;
        public IEnumerable<IObjectStateUDT> ObjectStateUDTCandidates
        {
            get
            {
                 if (_objStateCandidates != null)
                {
                    return _objStateCandidates;
                }

                _objStateCandidates = new HashSet<IObjectStateUDT>();
                foreach (var candidate in UDTFieldCandidates.Where(udt => udt.CanBeObjectStateUDT))
                {
                    _objStateCandidates.Add(new ObjectStateUDT(candidate));
                }

                _objStateCandidates.Add(_newObjectStateUDT);
                return _objStateCandidates;
            }
        }
    }
}
