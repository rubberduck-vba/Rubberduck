using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences
{
    public class ReplacePrivateUDTMemberReferencesModel : IRefactoringModel
    {
        private Dictionary<(VariableDeclaration, Declaration), PrivateUDTMemberReferenceReplacementExpressions> _udtTargets 
            = new Dictionary<(VariableDeclaration, Declaration), PrivateUDTMemberReferenceReplacementExpressions>();

        private Dictionary<VariableDeclaration, UserDefinedTypeInstance> _fieldToUserDefinedTypeInstance;

        public ReplacePrivateUDTMemberReferencesModel(Dictionary<VariableDeclaration, UserDefinedTypeInstance> fieldToUserDefinedTypeInstance, IEnumerable<Declaration> userDefinedTypeMembers)
        {
            _fieldToUserDefinedTypeInstance = fieldToUserDefinedTypeInstance;
            _udtMembers = userDefinedTypeMembers.ToList();
        }

        public IReadOnlyCollection<VariableDeclaration> Targets => _fieldToUserDefinedTypeInstance.Keys;

        private List<Declaration> _udtMembers;
        public IReadOnlyCollection<Declaration> UDTMembers => _udtMembers;

        public UserDefinedTypeInstance UserDefinedTypeInstance(VariableDeclaration field) 
            => _fieldToUserDefinedTypeInstance[field];

        public void AssignUDTMemberReferenceExpressions(VariableDeclaration field, Declaration udtMember, PrivateUDTMemberReferenceReplacementExpressions expressions)
        {
            _udtTargets.Add((field,udtMember), expressions);
        }

        public (bool HasValue, string Expression) LocalReferenceExpression(VariableDeclaration field, Declaration udtMember)
        {
            if (_udtTargets.TryGetValue((field, udtMember), out var result))
            {
                return (true, result.LocalReferenceExpression);
            }
            return (false, null);
        }
    }
}
