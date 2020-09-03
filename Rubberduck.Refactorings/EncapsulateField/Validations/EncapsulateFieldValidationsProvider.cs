using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public enum NameValidators
    {
        Default,
        UserDefinedType,
        UserDefinedTypeMember,
        UserDefinedTypeMemberArray
    }

    public class EncapsulateFieldValidationsProvider
    {
        private static readonly Dictionary<NameValidators, IValidateVBAIdentifiers> _nameOnlyValidators = new Dictionary<NameValidators, IValidateVBAIdentifiers>()
        {
            [NameValidators.Default] = new IdentifierOnlyValidator(DeclarationType.Variable, false),
            [NameValidators.UserDefinedType] = new IdentifierOnlyValidator(DeclarationType.UserDefinedType, false),
            [NameValidators.UserDefinedTypeMember] = new IdentifierOnlyValidator(DeclarationType.UserDefinedTypeMember, false),
            [NameValidators.UserDefinedTypeMemberArray] = new IdentifierOnlyValidator(DeclarationType.UserDefinedTypeMember, true),
        };

        public EncapsulateFieldValidationsProvider()
        {}

        public static IValidateVBAIdentifiers NameOnlyValidator(NameValidators validatorType)
            => _nameOnlyValidators[validatorType];
    }
}
