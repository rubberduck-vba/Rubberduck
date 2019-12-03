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
        private readonly Func<EncapsulateFieldModel, string> _previewFunc;
        private QualifiedModuleName _targetQMN;

        private bool _useNewStructure;

        private IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> _udtFieldToUdtDeclarationMap = new Dictionary<Declaration, (Declaration, IEnumerable<Declaration>)>();

        public EncapsulateFieldModel(Declaration target, IDeclarationFinderProvider declarationFinderProvider, /*IEnumerable<Declaration> allMemberFields, IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> udtFieldToUdtDeclarationMap,*/ IIndenter indenter, IEncapsulateFieldNamesValidator validator, Func<EncapsulateFieldModel, string> previewFunc)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _indenter = indenter;
            _validator = validator;
            _previewFunc = previewFunc;
            _targetQMN = target.QualifiedModuleName;

            _useNewStructure = File.Exists("C:\\Users\\Brian\\Documents\\UseNewUDTStructure.txt");

            CandidateFactory = new EncapsulationCandidateFactory(declarationFinderProvider, _validator);
            var encapsulationCandidateFields = _declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN)
                .Where(v => v.IsMemberVariable() && !v.IsWithEvents);

            var candidates = CandidateFactory.CreateEncapsulationCandidates(encapsulationCandidateFields);

            FieldCandidates.AddRange(candidates);

            AllEncapsulationCandidates = Enumerable.Empty<IEncapsulateFieldCandidate>()
                .Concat(FieldCandidates)
                .Concat(UDTFieldCandidates.SelectMany(c => c.Members)).ToList();

            EncapsulationStrategy = new EncapsulateWithBackingFields(_targetQMN, _indenter, _validator);
            this[target].EncapsulateFlag = true;

        }

        public EncapsulationCandidateFactory CandidateFactory { private set; get; }

        public List<IEncapsulateFieldCandidate> FieldCandidates { set; get; } = new List<IEncapsulateFieldCandidate>();

        public IEnumerable<IEncapsulatedUserDefinedTypeField> UDTFieldCandidates => FieldCandidates.Where(v => v is IEncapsulatedUserDefinedTypeField).Cast<IEncapsulatedUserDefinedTypeField>();

        private List<IEncapsulateFieldCandidate> AllEncapsulationCandidates { get; } = new List<IEncapsulateFieldCandidate>();

        public IEnumerable<IEncapsulateFieldCandidate> FlaggedEncapsulationFields
            => AllEncapsulationCandidates.Where(v => v.EncapsulationAttributes.EncapsulateFlag);

        public IEnumerable<IEncapsulateFieldCandidate> EncapsulationFields
            => AllEncapsulationCandidates;

        public IEncapsulateFieldCandidate this[string encapsulatedFieldIdentifier]
            => AllEncapsulationCandidates.Where(c => c.TargetID.Equals(encapsulatedFieldIdentifier)).Single();

        public IEncapsulateFieldCandidate this[Declaration fieldDeclaration]
            => AllEncapsulationCandidates.Where(c => c.Declaration == fieldDeclaration).Single();

        public IEncapsulateFieldStrategy EncapsulationStrategy { set; get; }

        public bool EncapsulateWithUDT
        {
            set
            {
                if (value && EncapsulationStrategy is EncapsulateWithBackingUserDefinedType) { return; }

                if (!value && EncapsulationStrategy is EncapsulateWithBackingFields) { return; }

                if (value)
                {
                    EncapsulationStrategy = new EncapsulateWithBackingUserDefinedType(_targetQMN, _indenter, _validator)
                    {
                        StateUDTField = CandidateFactory.CreateStateUDTField(_targetQMN)
                    };
                }
                else
                {
                    EncapsulationStrategy = new EncapsulateWithBackingFields(_targetQMN, _indenter, _validator);
                }
            }

            get => EncapsulationStrategy is EncapsulateWithBackingUserDefinedType;
        }

        public string PreviewRefactoring()
        {
            return _previewFunc(this);
        }

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
