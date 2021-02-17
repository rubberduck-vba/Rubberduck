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
            var udtInstanceToLowestLeafMembers = new Dictionary<VariableDeclaration, IEnumerable<Declaration>>();
            foreach (var target in targets)
            {
                udtInstanceToLowestLeafMembers = Load(udtInstanceToLowestLeafMembers, target);
            }

            var fieldToUDTInstance = new Dictionary<VariableDeclaration, UserDefinedTypeInstance>();
            foreach (var field in udtInstanceToLowestLeafMembers.Keys)
            {
                fieldToUDTInstance.Add(field, new UserDefinedTypeInstance(field, udtInstanceToLowestLeafMembers[field]));
            }

            var allUDTMembers = udtInstanceToLowestLeafMembers.Values.SelectMany(v => v);
            return new ReplacePrivateUDTMemberReferencesModel(fieldToUDTInstance, allUDTMembers.Distinct());
        }

        private Dictionary<VariableDeclaration, IEnumerable<Declaration>> Load(Dictionary<VariableDeclaration, IEnumerable<Declaration>> udtInstanceToLowestLeafMembers, Declaration target)
        {
            var udtMembers = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.UserDefinedTypeMember)
                .Where(udtm => udtm.ParentDeclaration == target.AsTypeDeclaration).ToList();

            var guard = 0;
            while (UdtMembersContainsUDTInstances(udtMembers))
            {
                if (guard++ >= 20)
                {
                    //TODO: Better exception
                    throw new System.IndexOutOfRangeException();
                }

                udtMembers = LowestLeafUdtMembers(udtMembers);
            }


            udtInstanceToLowestLeafMembers.Add(target as VariableDeclaration, udtMembers);
            return udtInstanceToLowestLeafMembers;
        }

        private List<Declaration> LowestLeafUdtMembers(List<Declaration> udtMembers)
        {
            var UDTsInMembers = udtMembers.Where(d => d.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ?? false);
            var childrenUDTMembers = new List<Declaration>();
            foreach (var udt in UDTsInMembers)
            {
                var children = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.UserDefinedTypeMember)
                    .Where(udtm => udtm.ParentDeclaration == udt.AsTypeDeclaration);

                childrenUDTMembers.AddRange(children);
            }

            udtMembers.RemoveAll(d => UDTsInMembers.Contains(d));
            udtMembers.AddRange(childrenUDTMembers);
            return udtMembers;
        }

        private bool UdtMembersContainsUDTInstances(List<Declaration> udtMembers) 
            => udtMembers.Any(d => d.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ?? false);
    }
}
