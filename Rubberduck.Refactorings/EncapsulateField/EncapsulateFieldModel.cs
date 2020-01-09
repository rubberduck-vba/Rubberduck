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
        private IDeclarationFinderProvider _declarationFinderProvider;
        private IEncapsulateFieldValidationsProvider _validatorProvider;
        private IObjectStateUDT _newObjectStateUDT;

        private IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> _udtFieldToUdtDeclarationMap = new Dictionary<Declaration, (Declaration, IEnumerable<Declaration>)>();

        public EncapsulateFieldModel(Declaration target, IEnumerable<IEncapsulatableField> candidates, IObjectStateUDT stateUDTField, Func<EncapsulateFieldModel, string> previewDelegate, IDeclarationFinderProvider declarationFinderProvider, IEncapsulateFieldValidationsProvider validatorProvider) // IEncapsulateFieldValidator validator)
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
            get => _encapsulationFieldStategy;
            set
            {
                _encapsulationFieldStategy = value;
                if (_encapsulationFieldStategy == EncapsulateFieldStrategy.UseBackingFields)
                {
                    foreach (var candidate in EncapsulationCandidates)
                    {
                        candidate.ConflictFinder = _validatorProvider.ConflictDetector(_encapsulationFieldStategy, _declarationFinderProvider);// ConflictDetector;
                        switch (candidate)
                        {
                            case IUserDefinedTypeCandidate udt:
                                candidate.NameValidator = _validatorProvider.NameOnlyValidator(NameValidators.UserDefinedType);
                                break;
                            case IUserDefinedTypeMemberCandidate udtm:
                                candidate.NameValidator = candidate.Declaration.IsArray
                                    ? _validatorProvider.NameOnlyValidator(NameValidators.UserDefinedTypeMemberArray)
                                    : _validatorProvider.NameOnlyValidator(NameValidators.UserDefinedTypeMember);
                                break;
                            default:
                                candidate.NameValidator = _validatorProvider.NameOnlyValidator(NameValidators.Default);
                                break;
                        }
                    }
                }
                else
                {
                    foreach (var candidate in EncapsulationCandidates)
                    {
                        candidate.ConflictFinder = _validatorProvider.ConflictDetector(_encapsulationFieldStategy, _declarationFinderProvider);// ConflictDetector;
                        candidate.NameValidator = candidate.Declaration.IsArray
                            ? _validatorProvider.NameOnlyValidator(NameValidators.UserDefinedTypeMemberArray)
                            : _validatorProvider.NameOnlyValidator(NameValidators.UserDefinedTypeMember);
                    }
                }
            }
        } 

        public IEncapsulateFieldValidationsProvider ValidationsProvider => _validatorProvider;

        public List<IEncapsulatableField> EncapsulationCandidates { set; get; } = new List<IEncapsulatableField>();

        public IEnumerable<IEncapsulatableField> SelectedFieldCandidates
            => EncapsulationCandidates.Where(v => v.EncapsulateFlag);

        public IEnumerable<IUserDefinedTypeCandidate> UDTFieldCandidates 
            => EncapsulationCandidates
                    .Where(v => v is IUserDefinedTypeCandidate)
                    .Cast<IUserDefinedTypeCandidate>();

        public IEnumerable<IUserDefinedTypeCandidate> SelectedUDTFieldCandidates 
            => SelectedFieldCandidates
                    .Where(v => v is IUserDefinedTypeCandidate)
                    .Cast<IUserDefinedTypeCandidate>();

        public IEncapsulatableField this[string encapsulatedFieldTargetID]
            => EncapsulationCandidates.Where(c => c.TargetID.Equals(encapsulatedFieldTargetID)).Single();

        public IEncapsulatableField this[Declaration fieldDeclaration]
            => EncapsulationCandidates.Where(c => c.Declaration == fieldDeclaration).Single();
        
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
