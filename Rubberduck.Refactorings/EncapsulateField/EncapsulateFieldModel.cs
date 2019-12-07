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

        public EncapsulateFieldModel(Declaration target, IDeclarationFinderProvider declarationFinderProvider, IIndenter indenter, IEncapsulateFieldNamesValidator validator, Func<EncapsulateFieldModel, string> previewDelegate)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _indenter = indenter;
            _validator = validator;
            _previewDelegate = previewDelegate;
            _targetQMN = target.QualifiedModuleName;

            _useNewStructure = File.Exists("C:\\Users\\Brian\\Documents\\UseNewUDTStructure.txt");

            CandidateFactory = new EncapsulationCandidateFactory(declarationFinderProvider, _validator);
            var encapsulationCandidateFields = _declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN)
                .Where(v => v.IsMemberVariable() && !v.IsWithEvents);

            var candidates = CandidateFactory.CreateEncapsulationCandidates(encapsulationCandidateFields);

            FieldCandidates.AddRange(candidates);

            //EncapsulationStrategy = new EncapsulateWithBackingFields(_targetQMN, _indenter, _validator);
            this[target].EncapsulateFlag = true;

        }

        public EncapsulationCandidateFactory CandidateFactory { private set; get; }

        public List<IEncapsulateFieldCandidate> FieldCandidates { set; get; } = new List<IEncapsulateFieldCandidate>();

        public IEnumerable<IEncapsulateFieldCandidate> FlaggedFieldCandidates
            => FieldCandidates.Where(v => v.EncapsulateFlag);

        public IEnumerable<IEncapsulatedUserDefinedTypeField> UDTFieldCandidates 
            => FieldCandidates
                    .Where(v => v is IEncapsulatedUserDefinedTypeField)
                    .Cast<IEncapsulatedUserDefinedTypeField>();

        public IEnumerable<IEncapsulatedUserDefinedTypeField> FlaggedUDTFieldCandidates 
            => FlaggedFieldCandidates
                    .Where(v => v is IEncapsulatedUserDefinedTypeField)
                    .Cast<IEncapsulatedUserDefinedTypeField>();

        public IEnumerable<IEncapsulateFieldCandidate> FlaggedEncapsulationFields
            => FlaggedFieldCandidates;

        public IEncapsulateFieldCandidate this[string encapsulatedFieldTargetID]
            => FieldCandidates.Where(c => c.TargetID.Equals(encapsulatedFieldTargetID)).Single();

        public IEncapsulateFieldCandidate this[Declaration fieldDeclaration]
            => FieldCandidates.Where(c => c.Declaration == fieldDeclaration).Single();

        public IEncapsulateFieldStrategy EncapsulationStrategy
        {
            get
            {
                if (EncapsulateWithUDT)
                {
                    return new EncapsulateWithBackingUserDefinedType(_targetQMN, _indenter, _validator)
                    {
                        StateUDTField = CandidateFactory.CreateStateUDTField(_targetQMN)
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
