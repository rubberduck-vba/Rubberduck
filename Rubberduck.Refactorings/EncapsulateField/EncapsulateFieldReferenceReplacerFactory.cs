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
        private IUDTMemberReferenceProvider _userDefinedTypeInstanceProvider;
        public EncapsulateFieldReferenceReplacerFactory(IDeclarationFinderProvider declarationFinderProvider,
            IPropertyAttributeSetsGenerator propertyAttributeSetsGenerator,
            IUDTMemberReferenceProvider userDefinedTypeInstanceProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _propertyAttributeSetsGenerator = propertyAttributeSetsGenerator;
            _userDefinedTypeInstanceProvider = userDefinedTypeInstanceProvider;
        }

        public IEncapsulateFieldReferenceReplacer Create()
        {
            return new EncapsulateFieldReferenceReplacer(_declarationFinderProvider,
                _propertyAttributeSetsGenerator,
                _userDefinedTypeInstanceProvider);
        }
    }
}
