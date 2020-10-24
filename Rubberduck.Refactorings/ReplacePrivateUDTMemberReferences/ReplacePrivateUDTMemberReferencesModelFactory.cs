using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings
{
    public interface IReplacePrivateUDTMemberReferencesModelFactory
    {
        ReplacePrivateUDTMemberReferencesModel Create(IEnumerable<VariableDeclaration> targets );
    }

    public class ReplacePrivateUDTMemberReferencesModelFactory : IReplacePrivateUDTMemberReferencesModelFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        public ReplacePrivateUDTMemberReferencesModelFactory(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public ReplacePrivateUDTMemberReferencesModel Create(IEnumerable<VariableDeclaration> targets)
        {
            var allUDTMembers = new List<Declaration>();
            var fieldsToUDTMembers = new Dictionary<VariableDeclaration, IEnumerable<Declaration>>();
            foreach (var target in targets)
            {
                var udtMembers = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.UserDefinedTypeMember)
                    .Where(udtm => udtm.ParentDeclaration == target.AsTypeDeclaration);

                allUDTMembers.AddRange(udtMembers);
                fieldsToUDTMembers.Add(target as VariableDeclaration, udtMembers);
            }

            var fieldToUDTInstance = new Dictionary<VariableDeclaration, UserDefinedTypeInstance>();
            foreach (var fieldToUDTMembers in fieldsToUDTMembers)
            {
                fieldToUDTInstance.Add(fieldToUDTMembers.Key, new UserDefinedTypeInstance(fieldToUDTMembers.Key, fieldToUDTMembers.Value));
            }

            return new ReplacePrivateUDTMemberReferencesModel(fieldToUDTInstance, allUDTMembers.Distinct());
        }
    }
}
