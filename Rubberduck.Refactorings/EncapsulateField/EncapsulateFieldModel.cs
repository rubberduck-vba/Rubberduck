using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldModel : IRefactoringModel
    {
        private readonly Func<EncapsulateFieldModel, string> _previewDelegate;
        private QualifiedModuleName _targetQMN;
        private IEncapsulateFieldNamesValidator _validator;

        private IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> _udtFieldToUdtDeclarationMap = new Dictionary<Declaration, (Declaration, IEnumerable<Declaration>)>();

        public EncapsulateFieldModel(Declaration target, IEnumerable<IEncapsulateFieldCandidate> candidates, IObjectStateUDT stateUDTField, Func<EncapsulateFieldModel, string> previewDelegate, IEncapsulateFieldNamesValidator validator)
        {
            _previewDelegate = previewDelegate;
            _targetQMN = target.QualifiedModuleName;
            _validator = validator;

            EncapsulationCandidates = candidates.ToList();
            //StateUDTField = stateUDTField;
        }

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

        public bool EncapsulateWithUDT { set; get; }

        private IObjectStateUDT _stateUDTField;
        public IObjectStateUDT StateUDTField
        {
            set
            {
                _stateUDTField = value;
            }
            get
            {
                if (_stateUDTField != null)
                {
                    return _stateUDTField;
                }

                if (!EncapsulateWithUDT) { return null; }
                var stateUDT = EncapsulationCandidates.Where(sfc => sfc is IUserDefinedTypeCandidate udt
                        && udt.IsObjectStateUDT).Select(sfc => sfc as IUserDefinedTypeCandidate).FirstOrDefault();

                _stateUDTField = stateUDT != null
                    ? new ObjectStateUDT(stateUDT)
                    : null;

                return _stateUDTField;
            }
        }

        public string PreviewRefactoring() => _previewDelegate(this);
    }
}
