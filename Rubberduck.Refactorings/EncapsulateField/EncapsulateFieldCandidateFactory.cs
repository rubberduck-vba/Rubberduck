using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField;
using System.Collections.Generic;
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
        private readonly ICodeBuilder _codeBuilder;

        private readonly IValidateVBAIdentifiers _defaultNameValidator;
        private readonly IValidateVBAIdentifiers _udtNameValidator;
        private readonly IValidateVBAIdentifiers _udtMemberNameValidator;
        private readonly IValidateVBAIdentifiers _udtMemberArrayNameValidator;

        public EncapsulateFieldCandidateFactory(IDeclarationFinderProvider declarationFinderProvider, ICodeBuilder codeBuilder)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _codeBuilder = codeBuilder;

            _defaultNameValidator = EncapsulateFieldValidationsProvider.NameOnlyValidator(NameValidators.Default);
            _udtNameValidator = EncapsulateFieldValidationsProvider.NameOnlyValidator(NameValidators.UserDefinedType);
            _udtMemberNameValidator = EncapsulateFieldValidationsProvider.NameOnlyValidator(NameValidators.UserDefinedTypeMember);
            _udtMemberArrayNameValidator = EncapsulateFieldValidationsProvider.NameOnlyValidator(NameValidators.UserDefinedTypeMemberArray);
        }

        public IEncapsulateFieldCandidate Create(Declaration target)
        {
            var candidate= CreateCandidate(target, _defaultNameValidator);
            return candidate;
        }

        private IEncapsulateFieldCandidate CreateCandidate(Declaration target, IValidateVBAIdentifiers validator)
        {
            if (target.IsUserDefinedType())
            {
                var udtField = new UserDefinedTypeCandidate(target, _udtNameValidator, _codeBuilder.BuildPropertyRhsParameterName) as IUserDefinedTypeCandidate;

                (Declaration udtDeclaration, IEnumerable<Declaration> udtMembers) = GetUDTAndMembersForField(udtField);

                foreach (var udtMemberDeclaration in udtMembers)
                {
                    var udtMemberValidator = _udtMemberNameValidator;
                    if (udtMemberDeclaration.IsArray)
                    {
                        udtMemberValidator = _udtMemberArrayNameValidator;
                    }
                    var candidateUDTMember = new UserDefinedTypeMemberCandidate(CreateCandidate(udtMemberDeclaration, udtMemberValidator), udtField, _codeBuilder.BuildPropertyRhsParameterName) as IUserDefinedTypeMemberCandidate;

                    udtField.AddMember(candidateUDTMember);
                }

                var udtVariablesOfSameType = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                    .Where(v => v.AsTypeDeclaration == udtDeclaration);

                udtField.CanBeObjectStateUDT = udtField.TypeDeclarationIsPrivate
                    && udtField.Declaration.HasPrivateAccessibility()
                    && udtVariablesOfSameType.Count() == 1;

                return udtField;
            }
            else if (target.IsArray)
            {
                var arrayCandidate = new ArrayCandidate(target, validator, _codeBuilder.BuildPropertyRhsParameterName);
                return arrayCandidate;
            }

            var candidate = new EncapsulateFieldCandidate(target, validator, _codeBuilder.BuildPropertyRhsParameterName);
            return candidate;
        }

        private (Declaration TypeDeclaration, IEnumerable<Declaration> Members) GetUDTAndMembersForField(IUserDefinedTypeCandidate udtField)
        {
            var userDefinedTypeDeclaration = udtField.Declaration.AsTypeDeclaration;

            var udtMembers = _declarationFinderProvider.DeclarationFinder
               .UserDeclarations(DeclarationType.UserDefinedTypeMember)
               .Where(utm => userDefinedTypeDeclaration == utm.ParentDeclaration);

            return (userDefinedTypeDeclaration, udtMembers);
        }
    }
}
