using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.Refactorings.CodeBlockInsert;
using Rubberduck.Refactorings.EncapsulateField;

namespace Rubberduck.Refactorings.EncapsulateFieldUseBackingField
{
    public class EncapsulateFieldUseBackingFieldModel : IRefactoringModel
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public EncapsulateFieldUseBackingFieldModel(IEnumerable<IEncapsulateFieldCandidate> candidates,
            IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;

            EncapsulationCandidates = candidates.ToList();

            ResetNewContent();
        }

        public void ResetNewContent()
        {
            NewContent = new Dictionary<NewContentType, List<string>>
            {
                { NewContentType.PostContentMessage, new List<string>() },
                { NewContentType.DeclarationBlock, new List<string>() },
                { NewContentType.CodeSectionBlock, new List<string>() },
                { NewContentType.TypeDeclarationBlock, new List<string>() }
            };
        }

        public IEncapsulateFieldConflictFinder ConflictFinder { set;  get; }

        public bool IncludeNewContentMarker { set; get; } = false;

        public List<IEncapsulateFieldCandidate> EncapsulationCandidates { get; }

        public IEnumerable<IEncapsulateFieldCandidate> SelectedFieldCandidates
            => EncapsulationCandidates.Where(v => v.EncapsulateFlag);

        public void AddContentBlock(NewContentType contentType, string block)
            => NewContent[contentType].Add(block);

        public Dictionary<NewContentType, List<string>> NewContent { set; get; }

        public QualifiedModuleName QualifiedModuleName => EncapsulationCandidates.First().QualifiedModuleName;
    }
}
