using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField.Strategies;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldModel : IRefactoringModel
    {
        private readonly IIndenter _indenter;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IEncapsulateFieldNamesValidator _validator;
        private readonly Func<EncapsulateFieldModel, string> _previewDelegate;
        private QualifiedModuleName _targetQMN;

        private bool _useNewStructure;

        private IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> _udtFieldToUdtDeclarationMap = new Dictionary<Declaration, (Declaration, IEnumerable<Declaration>)>();

        private EncapsulationCandidateFactory _fieldCandidateFactory;

        public EncapsulateFieldModel(Declaration target, IDeclarationFinderProvider declarationFinderProvider, IIndenter indenter, IEnumerable<IEncapsulateFieldCandidate> candidates, Func<EncapsulateFieldModel, string> previewDelegate)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _indenter = indenter;
            _previewDelegate = previewDelegate;
            _targetQMN = target.QualifiedModuleName;

            _useNewStructure = File.Exists("C:\\Users\\Brian\\Documents\\UseNewUDTStructure.txt");

            _validator = new EncapsulateFieldNamesValidator(_declarationFinderProvider); //, () => EncapsulationCandidates);

            EncapsulationCandidates = candidates.ToList();
            _fieldCandidateFactory = new EncapsulationCandidateFactory(declarationFinderProvider, _targetQMN, _validator);
            this[target].EncapsulateFlag = true;
        }

        public List<IEncapsulateFieldCandidate> EncapsulationCandidates { set; get; } = new List<IEncapsulateFieldCandidate>();

        public IEnumerable<IEncapsulateFieldCandidate> SelectedFieldCandidates
            => EncapsulationCandidates.Where(v => v.EncapsulateFlag);

        public IEnumerable<IEncapsulatedUserDefinedTypeField> UDTFieldCandidates 
            => EncapsulationCandidates
                    .Where(v => v is IEncapsulatedUserDefinedTypeField)
                    .Cast<IEncapsulatedUserDefinedTypeField>();

        public IEnumerable<IEncapsulatedUserDefinedTypeField> SelectedUDTFieldCandidates 
            => SelectedFieldCandidates
                    .Where(v => v is IEncapsulatedUserDefinedTypeField)
                    .Cast<IEncapsulatedUserDefinedTypeField>();

        public IEncapsulateFieldCandidate this[string encapsulatedFieldTargetID]
            => EncapsulationCandidates.Where(c => c.TargetID.Equals(encapsulatedFieldTargetID)).Single();

        public IEncapsulateFieldCandidate this[Declaration fieldDeclaration]
            => EncapsulationCandidates.Where(c => c.Declaration == fieldDeclaration).Single();

        public IEncapsulateFieldStrategy EncapsulationStrategy
        {
            get
            {
                if (EncapsulateWithUDT)
                {
                    return new EncapsulateWithBackingUserDefinedType(_targetQMN, _indenter, _validator)
                    {
                        StateUDTField = _fieldCandidateFactory.CreateStateUDTField()
                    };
                }
                return new EncapsulateWithBackingFields(_targetQMN, _indenter, _validator);
            }
        }

        public bool EncapsulateWithUDT { set; get; }

        public string PreviewRefactoring() => _previewDelegate(this);

        public int? CodeSectionStartIndex
        {
            get
            {
                var moduleMembers = _declarationFinderProvider.DeclarationFinder
                        .Members(_targetQMN).Where(m => m.IsMember());

                int? codeSectionStartIndex
                    = moduleMembers.OrderBy(c => c.Selection)
                                .FirstOrDefault()?.Context.Start.TokenIndex ?? null;

                return codeSectionStartIndex;
            }
        }
    }
}
