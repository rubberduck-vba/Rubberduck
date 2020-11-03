using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.ModifyUserDefinedType
{
    public class ModifyUserDefinedTypeModel : IRefactoringModel
    {
        private List<(Declaration, string)> _newMembers;
        private List<Declaration> _membersToRemove;

        public ModifyUserDefinedTypeModel(Declaration target)
        {
            if (!target.DeclarationType.HasFlag(DeclarationType.UserDefinedType))
            {
                throw new ArgumentException();
            }

            Target = target;
            _newMembers = new List<(Declaration, string)>();
            _membersToRemove = new List<Declaration>();
            InsertionIndex = (Target.Context as VBAParser.UdtDeclarationContext).END_TYPE().Symbol.TokenIndex - 1;
        }

        public Declaration Target { get; }

        public int InsertionIndex { get; }

        public void AddNewMemberPrototype(Declaration prototype, string memberIdentifier)
        {
            if (!IsValidPrototypeDeclarationType(prototype.DeclarationType))
            {
                throw new ArgumentException("Invalid prototype DeclarationType");
            }
            _newMembers.Add((prototype, memberIdentifier));
        }

        public void RemoveMember(Declaration member)
        {
            if (!member.DeclarationType.HasFlag(DeclarationType.UserDefinedTypeMember))
            {
                throw new ArgumentException();
            }
            _membersToRemove.Add(member);
        }

        public IEnumerable<(Declaration, string)> MembersToAdd => _newMembers;

        public IEnumerable<Declaration> MembersToRemove => _membersToRemove;

        private static bool IsValidPrototypeDeclarationType(DeclarationType declarationType)
        {
            return declarationType.HasFlag(DeclarationType.Variable)
                || declarationType.HasFlag(DeclarationType.UserDefinedTypeMember)
                || declarationType.HasFlag(DeclarationType.Constant)
                || declarationType.HasFlag(DeclarationType.Function);
        }
    }
}
