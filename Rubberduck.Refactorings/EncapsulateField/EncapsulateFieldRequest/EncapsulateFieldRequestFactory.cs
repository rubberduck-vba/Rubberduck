using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldRequestFactory
    {
        EncapsulateFieldRequest Create(VariableDeclaration target, bool isReadOnly = false, string propertyIdentifier = null);
    }

    public class EncapsulateFieldRequestFactory : IEncapsulateFieldRequestFactory
    {
        public EncapsulateFieldRequest Create(VariableDeclaration target, bool isReadOnly = false, string propertyIdentifier = null)
        {
            return new EncapsulateFieldRequest(target, isReadOnly, propertyIdentifier);
        }
    }
}
