using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldModel : IRefactoringModel
    {
        private readonly Func<EncapsulateFieldModel, string> _previewDelegate;
        private QualifiedModuleName _targetQMN;
        //private IValidateEncapsulateFieldNames _validator;
        private IObjectStateUDT _newObjectStateUDT;

        private IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> _udtFieldToUdtDeclarationMap = new Dictionary<Declaration, (Declaration, IEnumerable<Declaration>)>();

        public EncapsulateFieldModel(Declaration target, IEnumerable<IEncapsulateFieldCandidate> candidates, IObjectStateUDT stateUDTField, Func<EncapsulateFieldModel, string> previewDelegate, IEncapsulateFieldValidator validator)
        {
            _previewDelegate = previewDelegate;
            _targetQMN = target.QualifiedModuleName;
            _newObjectStateUDT = stateUDTField;

            Validator = validator;
            EncapsulationCandidates = candidates.ToList();
            ConvertFieldsToUDTMembers = false;
        }

        public IEncapsulateFieldValidator Validator {set; get;}

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

        private bool _convertFieldsToUDTMembers;
        public bool ConvertFieldsToUDTMembers
        {
            set
            {
                _convertFieldsToUDTMembers = value;
                SetFieldsToUDTMemberFlags(value);
            }
            get => _convertFieldsToUDTMembers;
        }

        private IObjectStateUDT _activeObjectStateUDT;
        public IObjectStateUDT StateUDTField
        {
            set => _activeObjectStateUDT = value;
            get => _activeObjectStateUDT ?? _newObjectStateUDT;
        }

        private void SetFieldsToUDTMemberFlags(bool value)
        {
            foreach (var candidate in EncapsulationCandidates)
            {
                candidate.ConvertFieldToUDTMember = value;
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
