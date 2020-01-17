using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IValidateVBAIdentifiers
    {
        bool IsValidVBAIdentifier(string identifier, out string errorMessage);
    }

    public class IdentifierOnlyValidator : IValidateVBAIdentifiers
    {
        private DeclarationType _declarationType;
        private bool _isArray;
        public IdentifierOnlyValidator(DeclarationType declarationType, bool isArray = false)
        {
            _declarationType = declarationType;
            _isArray = isArray;
        }

        public bool IsValidVBAIdentifier(string identifier, out string errorMessage)
            => !VBAIdentifierValidator.TryMatchInvalidIdentifierCriteria(identifier, _declarationType, out errorMessage, _isArray);
    }
}
