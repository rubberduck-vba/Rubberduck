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
        private const string DEFAULT_ENCAPSULATION_UDT_IDENTIFIER = "This_Type";
        private const string DEFAULT_ENCAPSULATION_UDT_FIELD_IDENTIFIER = "this";

        private readonly IIndenter _indenter;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IEncapsulateFieldNamesValidator _validator;
        private readonly Func<EncapsulateFieldModel, string> _previewFunc;
        private QualifiedModuleName _targetQMN;

        private bool _useNewStructure;

        private IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> _udtFieldToUdtDeclarationMap = new Dictionary<Declaration, (Declaration, IEnumerable<Declaration>)>();

        private Dictionary<string, IEncapsulateFieldCandidate> _candidates = new Dictionary<string, IEncapsulateFieldCandidate>();

        public EncapsulateFieldModel(Declaration target, IDeclarationFinderProvider declarationFinderProvider, /*IEnumerable<Declaration> allMemberFields, IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> udtFieldToUdtDeclarationMap,*/ IIndenter indenter, IEncapsulateFieldNamesValidator validator, Func<EncapsulateFieldModel, string> previewFunc)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _indenter = indenter;
            _validator = validator;
            _previewFunc = previewFunc;
            _targetQMN = target.QualifiedModuleName;

            _useNewStructure = File.Exists("C:\\Users\\Brian\\Documents\\UseNewUDTStructure.txt");

            //Maybe this should be passed in
            EncapsulationStrategy = new EncapsulateWithBackingFields(_targetQMN, _indenter, _declarationFinderProvider, _validator);
            _candidates = EncapsulationStrategy.FlattenedTargetIDToCandidateMapping;
            this[target].EncapsulateFlag = true;

        }

        public IEnumerable<IEncapsulateFieldCandidate> FlaggedEncapsulationFields
            => _candidates.Values.Where(v => v.EncapsulationAttributes.EncapsulateFlag);

        public IEnumerable<IEncapsulateFieldCandidate> EncapsulationFields
            => _candidates.Values;

        public IEncapsulateFieldCandidate this[string encapsulatedFieldIdentifier]
            => _candidates[encapsulatedFieldIdentifier];

        public IEncapsulateFieldCandidate this[Declaration fieldDeclaration]
            => _candidates.Values.Where(efd => efd.Declaration.Equals(fieldDeclaration)).Select(encapsulatedField => encapsulatedField).Single();

        public IEncapsulateFieldStrategy EncapsulationStrategy { set; get; }

        public bool EncapsulateWithUDT
        {
            set
            {
                if (EncapsulationStrategy is EncapsulateWithBackingUserDefinedType ebd) { return; }

                EncapsulationStrategy = value
                    ? new EncapsulateWithBackingUserDefinedType(_targetQMN, _indenter, _declarationFinderProvider, _validator) as IEncapsulateFieldStrategy
                    : new EncapsulateWithBackingFields(_targetQMN, _indenter, _declarationFinderProvider, _validator) as IEncapsulateFieldStrategy;
                //This probably should go - or be in the ctor
            }

            get => EncapsulationStrategy is EncapsulateWithBackingUserDefinedType;
        }

        //What to do with these as far as IStrategy goes - only here to support the view model
        public string EncapsulateWithUDT_TypeIdentifier { set; get; } = DEFAULT_ENCAPSULATION_UDT_IDENTIFIER;

        public string EncapsulateWithUDT_FieldName { set; get; } = DEFAULT_ENCAPSULATION_UDT_FIELD_IDENTIFIER;

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
