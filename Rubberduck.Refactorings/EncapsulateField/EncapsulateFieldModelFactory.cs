using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings
{
    public interface IEncapsulateFieldModelsFactory<T>
    {
        T Create(QualifiedModuleName qmn);
        T Create(IEnumerable<IEncapsulateFieldCandidate> candidates, IEnumerable<IObjectStateUDT> objectStateUDTCandidates);
    }

    public interface IEncapsulateFieldModelFactory
    {
        EncapsulateFieldModel Create(Declaration target);
    }

    public class EncapsulateFieldModelFactory : IEncapsulateFieldModelFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IEncapsulateFieldCandidateFactory _encapsulateFieldCandidateFactory;
        private readonly IEncapsulateFieldUseBackingUDTMemberModelFactory _useBackingUDTMemberModelFactory;
        private readonly IEncapsulateFieldUseBackingFieldModelFactory _useBackingFieldModelFactory;

        public EncapsulateFieldModelFactory(IDeclarationFinderProvider declarationFinderProvider, 
            IEncapsulateFieldCandidateFactory encapsulateFieldCandidateFactory,
            IEncapsulateFieldUseBackingUDTMemberModelFactory encapsulateFieldUseBackingUDTMemberModelFactory,
            IEncapsulateFieldUseBackingFieldModelFactory encapsulateFieldUseBackingFieldModelFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _encapsulateFieldCandidateFactory = encapsulateFieldCandidateFactory;
            _useBackingUDTMemberModelFactory = encapsulateFieldUseBackingUDTMemberModelFactory;
            _useBackingFieldModelFactory = encapsulateFieldUseBackingFieldModelFactory;
        }

        public EncapsulateFieldModel Create(Declaration target)
        {
            var fields = _declarationFinderProvider.DeclarationFinder
                .Members(target.QualifiedModuleName)
                .Where(v => v.IsMemberVariable() && !v.IsWithEvents);

            var candidates = fields.Select(fd => _encapsulateFieldCandidateFactory.Create(fd))
                .ToList();

            var objectStateUDTCandidates = candidates.Where(c => c is IUserDefinedTypeCandidate udt
                && udt.CanBeObjectStateUDT)
                .Select(udtc => new ObjectStateUDT(udtc as IUserDefinedTypeCandidate))
                .ToList();

            var initialStrategy = objectStateUDTCandidates
                .Any(os => os.AsTypeDeclaration.IdentifierName.StartsWith($"T{target.QualifiedModuleName.ComponentName}", System.StringComparison.InvariantCultureIgnoreCase))
                    ? EncapsulateFieldStrategy.ConvertFieldsToUDTMembers
                    : EncapsulateFieldStrategy.UseBackingFields;

            var selected = candidates.Single(c => c.Declaration == target);
            selected.EncapsulateFlag = true;

            var udtModel = _useBackingUDTMemberModelFactory.Create(candidates, objectStateUDTCandidates);
            var backingModel = _useBackingFieldModelFactory.Create(candidates, objectStateUDTCandidates);

            return new EncapsulateFieldModel(backingModel, udtModel)
            {
                EncapsulateFieldStrategy = initialStrategy
            };
        }
    }
}
