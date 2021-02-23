using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class UserDefinedTypeInstanceProvider
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly Dictionary<VariableDeclaration, UserDefinedTypeInstance> _fieldToUserDefinedTypeInstance;
        private readonly List<Declaration> _udtMembers;
        private readonly List<UserDefinedTypeInstance> _udtInstances = new List<UserDefinedTypeInstance>();
        public UserDefinedTypeInstanceProvider(IDeclarationFinderProvider declarationFinderProvider, IEnumerable<Declaration> udtInstances)
        {
            _declarationFinderProvider = declarationFinderProvider;
            var udtInstanceToLowestLeafMembers = new Dictionary<VariableDeclaration, IEnumerable<Declaration>>();
            foreach (var target in udtInstances)
            {
                if (target.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ?? false)
                {
                    udtInstanceToLowestLeafMembers = Load(udtInstanceToLowestLeafMembers, target);
                }
            }

            foreach (var field in udtInstanceToLowestLeafMembers.Keys)
            {
                _udtInstances.Add(new UserDefinedTypeInstance(field, udtInstanceToLowestLeafMembers[field]));
            }
            _udtMembers = udtInstanceToLowestLeafMembers.Values.SelectMany(v => v).ToList();

            _fieldToUserDefinedTypeInstance = _udtInstances.ToDictionary(u => u.InstanceField);
        }
        public IReadOnlyCollection<VariableDeclaration> Targets => _fieldToUserDefinedTypeInstance.Keys;

        public IReadOnlyCollection<Declaration> UDTMembers => _udtMembers;

        public UserDefinedTypeInstance this[VariableDeclaration field]
            => _fieldToUserDefinedTypeInstance[field];
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

                udtMembers = LowestLeafUdtMembers(udtMembers, _declarationFinderProvider);
            }


            udtInstanceToLowestLeafMembers.Add(target as VariableDeclaration, udtMembers);
            return udtInstanceToLowestLeafMembers;
        }
        private static List<Declaration> LowestLeafUdtMembers(List<Declaration> udtMembers, IDeclarationFinderProvider declarationFinderProvider)
        {
            var UDTsInMembers = udtMembers.Where(d => d.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ?? false);
            var childrenUDTMembers = new List<Declaration>();
            foreach (var udt in UDTsInMembers)
            {
                var children = declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.UserDefinedTypeMember)
                    .Where(udtm => udtm.ParentDeclaration == udt.AsTypeDeclaration);

                childrenUDTMembers.AddRange(children);
            }

            udtMembers.RemoveAll(d => UDTsInMembers.Contains(d));
            udtMembers.AddRange(childrenUDTMembers);
            return udtMembers;
        }
        private static bool UdtMembersContainsUDTInstances(List<Declaration> udtMembers)
            => udtMembers.Any(d => d.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ?? false);
    }

}
