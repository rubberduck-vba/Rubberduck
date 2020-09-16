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

                var udtVariablesOfSameType = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                    .Where(v => v.AsTypeDeclaration == udtField.Declaration.AsTypeDeclaration);

                udtField.IsObjectStateUDTCandidate = udtField.TypeDeclarationIsPrivate
                    && udtField.Declaration.HasPrivateAccessibility()
                    && udtVariablesOfSameType.Count() == 1;

                return udtField;
            }
            else if (target.IsArray)
            {
                var arrayCandidate = new ArrayCandidate(target);
                return arrayCandidate;
            }

            return new EncapsulateFieldCandidate(target);
        }
    }
}
