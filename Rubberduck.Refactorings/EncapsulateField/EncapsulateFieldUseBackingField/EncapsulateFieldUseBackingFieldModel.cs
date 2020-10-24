using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.Refactorings.EncapsulateField;

namespace Rubberduck.Refactorings.EncapsulateFieldUseBackingField
{
    public class EncapsulateFieldUseBackingFieldModel : IRefactoringModel
    {
        public EncapsulateFieldUseBackingFieldModel(IEnumerable<IEncapsulateFieldCandidate> candidates)
        {
            EncapsulationCandidates = candidates.ToList();
            if (EncapsulationCandidates.Any())
            {
                QualifiedModuleName = EncapsulationCandidates.First().QualifiedModuleName;
            }
        }

        public INewContentAggregator NewContentAggregator { set; get; }

        public IEncapsulateFieldConflictFinder ConflictFinder { set;  get; }

        public IReadOnlyCollection<IEncapsulateFieldCandidate> EncapsulationCandidates { get; }

        public IReadOnlyCollection<IEncapsulateFieldCandidate> SelectedFieldCandidates
            => EncapsulationCandidates.Where(c => c.EncapsulateFlag).ToList();

        public QualifiedModuleName QualifiedModuleName { get; } = new QualifiedModuleName();
    }
}
