using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldNamesValidator
    {
        bool HasValidEncapsulationAttributes(IEncapsulatedFieldDeclaration encapsulatedField);
        bool HasValidEncapsulationAttributes(IEnumerable<IEncapsulatedFieldDeclaration> newMembers);
        bool IsConflictingMemberName(string identifier, QualifiedModuleName qmn, DeclarationType declarationType);
    }

    public class EncapsulateFieldNamesValidator : IEncapsulateFieldNamesValidator
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        public EncapsulateFieldNamesValidator(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public bool HasValidEncapsulationAttributes(IEncapsulatedFieldDeclaration encapsulatedField)
        {
            var internalChecks = HasValidIdentifiers(encapsulatedField)
                && !HasInternalNameConflicts(encapsulatedField);

            var external = _declarationFinderProvider.DeclarationFinder
                .FindNewDeclarationNameConflicts(encapsulatedField.NewFieldName, encapsulatedField.Declaration);

            //if the collision is the target, then ignore
            if (external.Count() == 1 && external.First() == encapsulatedField.Declaration)
            {
                return internalChecks;
            }
            return internalChecks && !external.Any();
        }

        public bool HasValidEncapsulationAttributes(IEnumerable<IEncapsulatedFieldDeclaration> newMembers)
        {
            return newMembers.Any(nm => !HasValidIdentifiers(nm));
        }

        public bool IsConflictingMemberName(string memberName, QualifiedModuleName qmn, DeclarationType declarationType)
        {
            var members = _declarationFinderProvider.DeclarationFinder.Members(qmn);
            return members.Any(m => m.IdentifierName.Equals(memberName, StringComparison.CurrentCultureIgnoreCase));
        }

        private bool HasValidIdentifiers(IEncapsulatedFieldDeclaration encapsulatedField)
        {
            var attributes = encapsulatedField.EncapsulationAttributes;
            return VBAIdentifierValidator.IsValidIdentifier(attributes.NewFieldName, encapsulatedField.Declaration.DeclarationType)
                && VBAIdentifierValidator.IsValidIdentifier(attributes.PropertyName, encapsulatedField.Declaration.DeclarationType)
                && VBAIdentifierValidator.IsValidIdentifier(attributes.ParameterName, encapsulatedField.Declaration.DeclarationType);
        }

        private bool HasInternalNameConflicts(IEncapsulatedFieldDeclaration encapsulatedField)
        {
            var attributes = encapsulatedField.EncapsulationAttributes;
            return attributes.PropertyName.Equals(attributes.NewFieldName, StringComparison.InvariantCultureIgnoreCase)
                || attributes.PropertyName.Equals(attributes.ParameterName, StringComparison.InvariantCultureIgnoreCase)
                || attributes.NewFieldName.Equals(attributes.ParameterName, StringComparison.InvariantCultureIgnoreCase);
        }
    }
}
