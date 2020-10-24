using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.ReplaceDeclarationIdentifier
{
    public class ReplaceDeclarationIdentifierModel : IRefactoringModel
    {
        private List<(Declaration, string)> _targetNewNamePairs;

        public ReplaceDeclarationIdentifierModel(Declaration target, string newName)
            : this((target, newName)) { }

        public ReplaceDeclarationIdentifierModel(params (Declaration, string)[] targetNewNamePairs)
            : this(targetNewNamePairs.ToList()) { }

        public ReplaceDeclarationIdentifierModel(IEnumerable<(Declaration, string)> targetNewNamePairs)
        {
            _targetNewNamePairs = targetNewNamePairs.ToList();
        }

        public IReadOnlyCollection<(Declaration, string)> TargetNewNamePairs => _targetNewNamePairs;
    }
}
