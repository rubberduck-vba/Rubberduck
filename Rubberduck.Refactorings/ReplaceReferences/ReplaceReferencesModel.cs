using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.ReplaceReferences
{
    public class ReplaceReferencesModel :IRefactoringModel
    {
        public ReplaceReferencesModel()
        {}
        private Dictionary<IdentifierReference, string> _fieldTargets = new Dictionary<IdentifierReference, string>();

        public bool ModuleQualifyExternalReferences { set; get; } = false;

        public void AssignReferenceReplacementExpression(IdentifierReference fieldReference, string replacementIdentifier)
        {
            if (_fieldTargets.ContainsKey(fieldReference))
            {
                _fieldTargets[fieldReference] = replacementIdentifier;
                return;
            }
            _fieldTargets.Add(fieldReference, replacementIdentifier);
        }
        public IReadOnlyCollection<(IdentifierReference IdentifierReference, string NewName)> ReferenceReplacementPairs
            => _fieldTargets.Select(t => (t.Key, t.Value)).ToList();
    }
}
