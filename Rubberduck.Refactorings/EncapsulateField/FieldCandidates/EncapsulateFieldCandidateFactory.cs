using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField;
using System.Linq;

namespace Rubberduck.Refactorings
{
    public interface IEncapsulateFieldCandidateFactory
    {
        IEncapsulateFieldCandidate Create(Declaration target);
    }

    public class EncapsulateFieldCandidateFactory : IEncapsulateFieldCandidateFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public EncapsulateFieldCandidateFactory(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public IEncapsulateFieldCandidate Create(Declaration target)
        {
            if (target.IsUserDefinedType())
            {
                var udtField = new UserDefinedTypeCandidate(target) as IUserDefinedTypeCandidate;

                var udtMembers = _declarationFinderProvider.DeclarationFinder
                    .UserDeclarations(DeclarationType.UserDefinedTypeMember)
                    .Where(utm => udtField.Declaration.AsTypeDeclaration == utm.ParentDeclaration);

                foreach (var udtMemberDeclaration in udtMembers)
                {
                    var candidateUDTMember = new UserDefinedTypeMemberCandidate(Create(udtMemberDeclaration), udtField) as IUserDefinedTypeMemberCandidate;
                    udtField.AddMember(candidateUDTMember);
                }

                return udtField;
            }
            else if (target.IsArray)
            {
                return new ArrayCandidate(target);
            }

            return new EncapsulateFieldCandidate(target);
        }
    }
}
