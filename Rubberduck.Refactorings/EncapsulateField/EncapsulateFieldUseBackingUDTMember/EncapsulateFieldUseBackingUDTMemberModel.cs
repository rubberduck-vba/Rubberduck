using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.Refactorings.CodeBlockInsert;
using Rubberduck.Refactorings.EncapsulateField;
using System;

namespace Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember
{
    public class EncapsulateFieldUseBackingUDTMemberModel : IRefactoringModel
    {
        private readonly IObjectStateUDT _defaultObjectStateUDT;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly string _defaultObjectStateUDTTypeName;
        private readonly IObjectStateUDT _preExistingObjectStateUDT;

        private List<IConvertToUDTMember> _convertedFields;
        private List<IObjectStateUDT> _objStateCandidates;

        public EncapsulateFieldUseBackingUDTMemberModel(IEnumerable<IConvertToUDTMember> candidates,
            IObjectStateUDT defaultObjectStateUserDefinedType,
            IEnumerable<IObjectStateUDT> objectStateUserDefinedTypeCandidates,
            IDeclarationFinderProvider declarationFinderProvider)
        {
            _convertedFields = new List<IConvertToUDTMember>(candidates);
            _declarationFinderProvider = declarationFinderProvider;
            _defaultObjectStateUDT = defaultObjectStateUserDefinedType;

            QualifiedModuleName = candidates.First().QualifiedModuleName;
            _defaultObjectStateUDTTypeName = $"T{QualifiedModuleName.ComponentName}";

            _objStateCandidates = new List<IObjectStateUDT>();

            if (objectStateUserDefinedTypeCandidates.Any())
            {
                _objStateCandidates.AddRange(objectStateUserDefinedTypeCandidates.Distinct());

                _preExistingObjectStateUDT = objectStateUserDefinedTypeCandidates
                    .FirstOrDefault(os => os.AsTypeDeclaration.IdentifierName.StartsWith(_defaultObjectStateUDTTypeName, StringComparison.InvariantCultureIgnoreCase));

                if (_preExistingObjectStateUDT != null)
                {
                    HasPreExistingObjectStateUDT = true;
                    _defaultObjectStateUDT.IsSelected = false;
                    _preExistingObjectStateUDT.IsSelected = true;
                    _convertedFields.ForEach(c => c.ObjectStateUDT = _preExistingObjectStateUDT);
                }
            }

            _objStateCandidates.Add(_defaultObjectStateUDT);

            ResetNewContent();

            _convertedFields.ForEach(c => c.ObjectStateUDT = ObjectStateUDTField);
        }

        private void ResetNewContent()
        {
            NewContent = new Dictionary<NewContentType, List<string>>
            {
                { NewContentType.PostContentMessage, new List<string>() },
                { NewContentType.DeclarationBlock, new List<string>() },
                { NewContentType.CodeSectionBlock, new List<string>() },
                { NewContentType.TypeDeclarationBlock, new List<string>() }
            };
        }

        public bool HasPreExistingObjectStateUDT { get; }

        public IEncapsulateFieldConflictFinder ConflictFinder { set; get; }

        public bool IncludeNewContentMarker { set; get; } = false;

        public IReadOnlyCollection<IEncapsulateFieldCandidate> EncapsulationCandidates
            => _convertedFields.Cast<IEncapsulateFieldCandidate>().ToList();

        public IEnumerable<IEncapsulateFieldCandidate> SelectedFieldCandidates
            => EncapsulationCandidates.Where(v => v.EncapsulateFlag);

        public void AddContentBlock(NewContentType contentType, string block)
            => NewContent[contentType].Add(block);

        public Dictionary<NewContentType, List<string>> NewContent { set; get; }

        public QualifiedModuleName QualifiedModuleName { get; }

        public IEnumerable<IObjectStateUDT> ObjectStateUDTCandidates => _objStateCandidates;

        public IObjectStateUDT ObjectStateUDTField
        {
            get => _objStateCandidates.SingleOrDefault(os => os.IsSelected)
                ?? _defaultObjectStateUDT;

            set
            {
                if (value is null)
                {
                    _objStateCandidates.ForEach(osc => osc.IsSelected = (osc == _defaultObjectStateUDT));
                    return;
                }

                var matchingCandidate = _objStateCandidates
                    .SingleOrDefault(os => os.FieldIdentifier.Equals(value.FieldIdentifier))
                    ?? _defaultObjectStateUDT;

                _objStateCandidates.ForEach(osc => osc.IsSelected = (osc == matchingCandidate));

                _convertedFields.ForEach(cf => cf.ObjectStateUDT = matchingCandidate);
            }
        }
    }
}
