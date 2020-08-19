using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.MoveFieldsToUDT
{
    public class DeclareFieldsAsUDTMembersModel : IRefactoringModel
    {
        private Dictionary<Declaration, List<(VariableDeclaration Field, string UDTMemberIdentifier)>> _targets { get; } = new Dictionary<Declaration, List<(VariableDeclaration, string)>>(); 

        public DeclareFieldsAsUDTMembersModel()
        {}

        public IReadOnlyCollection<Declaration> UserDefinedTypeTargets => _targets.Keys;

        public IEnumerable<(VariableDeclaration Field, string userDefinedTypeMemberIdentifier)> this[Declaration udt] 
            => _targets[udt].Select(pr => (pr.Field, pr.UDTMemberIdentifier));

        public void AssignFieldToUserDefinedType(Declaration udt, VariableDeclaration field, string udtMemberIdentifierName = null)
        {
            if (!udt.DeclarationType.HasFlag(DeclarationType.UserDefinedType))
            {
                throw new ArgumentException();
            }

            if (!(_targets.TryGetValue(udt, out var memberPrototypes)))
            {
                _targets.Add(udt, new List<(VariableDeclaration, string)>());
            }
            else
            {
                var hasDuplicateMemberNames = memberPrototypes
                    .Select(pr => pr.UDTMemberIdentifier?.ToUpperInvariant() ?? pr.Field.IdentifierName)
                    .GroupBy(uc => uc).Any(g => g.Count() > 1);

                if (hasDuplicateMemberNames)
                {
                    throw new ArgumentException();
                }
            }

            _targets[udt].Add((field, udtMemberIdentifierName ?? field.IdentifierName));
        }
    }
}
