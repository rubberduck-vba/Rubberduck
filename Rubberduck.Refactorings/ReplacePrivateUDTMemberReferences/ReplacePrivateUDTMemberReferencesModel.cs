using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences
{
    public class ReplacePrivateUDTMemberReferencesModel : IRefactoringModel
    {
        private readonly Dictionary<IdentifierReference, string> _udtMemberReferenceReplacements = new Dictionary<IdentifierReference, string>();
        private readonly Dictionary<VariableDeclaration, UserDefinedTypeInstance> _fieldToUserDefinedTypeInstance;
        private readonly List<Declaration> _udtMembers;

        public ReplacePrivateUDTMemberReferencesModel(Dictionary<VariableDeclaration, UserDefinedTypeInstance> fieldToUserDefinedTypeInstance, IEnumerable<Declaration> userDefinedTypeMembers)
        {
            _fieldToUserDefinedTypeInstance = fieldToUserDefinedTypeInstance;
            _udtMembers = userDefinedTypeMembers.ToList();
        }

        public IReadOnlyCollection<VariableDeclaration> Targets => _fieldToUserDefinedTypeInstance.Keys;

        public IReadOnlyCollection<Declaration> UDTMembers => _udtMembers;

        public bool ModuleQualifyExternalReplacements { set; get; } = true;

        public void RegisterReferenceReplacementExpression(IdentifierReference rf, string expression)
        {
            if (ModuleQualifyExternalReplacements && rf.QualifiedModuleName != Targets.First().QualifiedModuleName)
            {
                expression = $"{Targets.First().QualifiedModuleName.ComponentName}.{expression}";
            }
            _udtMemberReferenceReplacements.Add(rf, expression);
        }

        public UserDefinedTypeInstance UserDefinedTypeInstance(VariableDeclaration field) 
            => _fieldToUserDefinedTypeInstance[field];

        public bool TryGetLocalReferenceExpression(IdentifierReference udtMemberRf, out string expression) 
            => _udtMemberReferenceReplacements.TryGetValue(udtMemberRf, out expression);
    }
}
