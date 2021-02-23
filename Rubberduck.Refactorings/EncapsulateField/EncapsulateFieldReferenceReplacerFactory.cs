using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;

namespace Rubberduck.Refactorings
{
    public interface IEncapsulateFieldReferenceReplacerFactory
    {
        IEncapsulateFieldReferenceReplacer Create();
    }
    public class EncapsulateFieldReferenceReplacerFactory : IEncapsulateFieldReferenceReplacerFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IPropertyAttributeSetsGenerator _propertyAttributeSetsGenerator;
        public EncapsulateFieldReferenceReplacerFactory(IDeclarationFinderProvider declarationFinderProvider,
            IPropertyAttributeSetsGenerator propertyAttributeSetsGenerator)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _propertyAttributeSetsGenerator = propertyAttributeSetsGenerator;
        }

        public IEncapsulateFieldReferenceReplacer Create()
        {
            return new EncapsulateFieldReferenceReplacer(_declarationFinderProvider,
                _propertyAttributeSetsGenerator);
        }
    }
}
