using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor;
using System;
using System.Linq;

namespace Rubberduck.Refactorings
{
    public interface IEncapsulateFieldCandidateFactory
    {
        IEncapsulateFieldCandidate CreateFieldCandidate(Declaration target);
        IEncapsulateFieldAsUDTMemberCandidate CreateUDTMemberCandidate(IEncapsulateFieldCandidate fieldCandidate, IObjectStateUDT defaultObjectStateField);
        IObjectStateUDT CreateDefaultObjectStateField(QualifiedModuleName qualifiedModuleName);
        IObjectStateUDT CreateObjectStateField(IUserDefinedTypeCandidate userDefinedTypeField);
    }

    public class EncapsulateFieldCandidateFactory : IEncapsulateFieldCandidateFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public EncapsulateFieldCandidateFactory(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public IEncapsulateFieldCandidate CreateFieldCandidate(Declaration target)
        {
            if (!target.IsUserDefinedType())
            {
                return new EncapsulateFieldCandidate(target);
            }

            var udtField = new UserDefinedTypeCandidate(target) as IUserDefinedTypeCandidate;

            var udtMembers = _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.UserDefinedTypeMember)
                .Where(utm => udtField.Declaration.AsTypeDeclaration == utm.ParentDeclaration);

            foreach (var udtMemberDeclaration in udtMembers)
            {
                var candidateUDTMember = new UserDefinedTypeMemberCandidate(CreateFieldCandidate(udtMemberDeclaration), udtField);
                udtField.AddMember(candidateUDTMember);
            }

            return udtField;
        }

        public IEncapsulateFieldAsUDTMemberCandidate CreateUDTMemberCandidate(IEncapsulateFieldCandidate fieldCandidate, IObjectStateUDT defaultObjectStateField) 
            => new EncapsulateFieldAsUDTMemberCandidate(fieldCandidate, defaultObjectStateField);

        public IObjectStateUDT CreateDefaultObjectStateField(QualifiedModuleName qualifiedModuleName) 
            => new ObjectStateFieldCandidate(qualifiedModuleName);

        public IObjectStateUDT CreateObjectStateField(IUserDefinedTypeCandidate userDefinedTypeField)
        {
            if ((userDefinedTypeField.Declaration.AsTypeDeclaration?.Accessibility ?? Accessibility.Implicit) != Accessibility.Private)
            {
                throw new ArgumentException();
            }

            return new ObjectStateFieldCandidate(userDefinedTypeField);
        }
    }
}
