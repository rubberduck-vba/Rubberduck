using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.ReplaceReferences
{
    public class ReplaceReferencesModel :IRefactoringModel
    {
        private Dictionary<IdentifierReference, string> _fieldTargets = new Dictionary<IdentifierReference, string>();

        public bool ModuleQualifyExternalReferences { set; get; } = false;

        public void AssignFieldReferenceReplacementExpression(IdentifierReference fieldReference, string replacementIdentifier)
        {
            if (_fieldTargets.ContainsKey(fieldReference))
            {
                _fieldTargets[fieldReference] = replacementIdentifier;
                return;
            }
            _fieldTargets.Add(fieldReference, replacementIdentifier);
        }
        public IReadOnlyList<(IdentifierReference IdentifierReference, string NewName)> FieldReferenceReplacementPairs
            => _fieldTargets.Select(t => (t.Key, t.Value)).ToList();
    }
}
