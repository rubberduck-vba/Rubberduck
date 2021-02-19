using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences;
using Rubberduck.Refactorings.ReplaceReferences;

namespace Rubberduck.Refactorings
{
    public interface IEncapsulateFieldReferenceReplacerFactory
    {
        IEncapsulateFieldReferenceReplacer Create();
    }
    public class EncapsulateFieldReferenceReplacerFactory : IEncapsulateFieldReferenceReplacerFactory
    {
        private readonly IReplacePrivateUDTMemberReferencesModelFactory _replacePrivateUDTMemberReferencesModelFactory;
        private readonly ICodeOnlyRefactoringAction<ReplacePrivateUDTMemberReferencesModel> _replacePrivateUDTMemberReferencesRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<ReplaceReferencesModel> _replaceReferencesRefactoringAction;
        private readonly IPropertyAttributeSetsGenerator _propertyAttributeSetsGenerator;
        public EncapsulateFieldReferenceReplacerFactory(IReplacePrivateUDTMemberReferencesModelFactory replacePrivateUDTMemberReferencesModelFactory,
            IEncapsulateFieldRefactoringActionsProvider refactoringActionsProvider,
            IPropertyAttributeSetsGenerator propertyAttributeSetsGenerator)
        {
            _replacePrivateUDTMemberReferencesModelFactory = replacePrivateUDTMemberReferencesModelFactory;
            _replacePrivateUDTMemberReferencesRefactoringAction = refactoringActionsProvider.ReplaceUDTMemberReferences;
            _replaceReferencesRefactoringAction = refactoringActionsProvider.ReplaceReferences;
            _propertyAttributeSetsGenerator = propertyAttributeSetsGenerator;
        }

        public IEncapsulateFieldReferenceReplacer Create()
        {
            return new EncapsulateFieldReferenceReplacer(
                _replacePrivateUDTMemberReferencesModelFactory,
                _replacePrivateUDTMemberReferencesRefactoringAction,
                _replaceReferencesRefactoringAction,
                _propertyAttributeSetsGenerator);
        }
    }
}
