using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldModel : IRefactoringModel
    {
        private readonly Func<EncapsulateFieldModel, string> _previewDelegate;
        private QualifiedModuleName _targetQMN;

        private bool _useNewStructure;

        private IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> _udtFieldToUdtDeclarationMap = new Dictionary<Declaration, (Declaration, IEnumerable<Declaration>)>();

        public EncapsulateFieldModel(Declaration target, IEnumerable<IEncapsulateFieldCandidate> candidates, IStateUDT stateUDTField, Func<EncapsulateFieldModel, string> previewDelegate)
        {
            _previewDelegate = previewDelegate;
            _targetQMN = target.QualifiedModuleName;

            _useNewStructure = File.Exists("C:\\Users\\Brian\\Documents\\UseNewUDTStructure.txt");

            EncapsulationCandidates = candidates.ToList();
            StateUDTField = stateUDTField;
            this[target].EncapsulateFlag = true;
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

        public bool HasSelectedMultipleUDTFieldsOfType(string asTypeName)
                => SelectedUDTFieldCandidates.Where(f => f.AsTypeName.Equals(asTypeName)).Count() > 1;

        public IEncapsulateFieldCandidate this[string encapsulatedFieldTargetID]
            => EncapsulationCandidates.Where(c => c.TargetID.Equals(encapsulatedFieldTargetID)).Single();

        public IEncapsulateFieldCandidate this[Declaration fieldDeclaration]
            => EncapsulationCandidates.Where(c => c.Declaration == fieldDeclaration).Single();

        public bool EncapsulateWithUDT { set; get; }

        public IStateUDT StateUDTField { set; get; }

        public string PreviewRefactoring() => _previewDelegate(this);
    }
}
