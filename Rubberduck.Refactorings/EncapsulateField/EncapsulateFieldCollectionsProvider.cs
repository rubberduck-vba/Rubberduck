using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldCollectionsProvider
    {
        IReadOnlyCollection<IEncapsulateFieldCandidate> EncapsulateFieldCandidates { get; }
        IReadOnlyCollection<IEncapsulateFieldAsUDTMemberCandidate> EncapsulateAsUserDefinedTypeMemberCandidates { get; }
        IReadOnlyCollection<IObjectStateUDT> ObjectStateUDTCandidates { get; }
    }

    /// <summary>
    /// EncapsulateFieldCollectionsProvider generates collections of IEncapsulateFieldCandidate
    /// instances, IEncapsulateFieldAsUDTMemberCandidate instances, and IObjectStateUDT instances.
    /// It provides these collection instances to the various objects of the EncapsulateFieldRefactoring.
    /// </summary>
    public class EncapsulateFieldCollectionsProvider : IEncapsulateFieldCollectionsProvider
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IEncapsulateFieldCandidateFactory _encapsulateFieldCandidateFactory;
        private readonly IObjectStateUserDefinedTypeFactory _objectStateUserDefinedTypeFactory;

        public EncapsulateFieldCollectionsProvider(
            IDeclarationFinderProvider declarationFinderProvider,
            IEncapsulateFieldCandidateFactory encapsulateFieldCandidateFactory,
            IObjectStateUserDefinedTypeFactory objectStateUserDefinedTypeFactory,
            QualifiedModuleName qualifiedModuleName)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _encapsulateFieldCandidateFactory = encapsulateFieldCandidateFactory;
            _objectStateUserDefinedTypeFactory = objectStateUserDefinedTypeFactory;

            EncapsulateFieldCandidates = _declarationFinderProvider.DeclarationFinder.Members(qualifiedModuleName, DeclarationType.Variable)
                .Where(v => v.ParentDeclaration is ModuleDeclaration
                    && !v.IsWithEvents)
                .Select(f => _encapsulateFieldCandidateFactory.Create(f))
                .ToList();

            ObjectStateUDTCandidates = LoadObjectStateUDTCandidates(EncapsulateFieldCandidates, _objectStateUserDefinedTypeFactory, qualifiedModuleName);

            EncapsulateAsUserDefinedTypeMemberCandidates = LoadAsUDTMemberCandidates(EncapsulateFieldCandidates, ObjectStateUDTCandidates);
        }

        public IReadOnlyCollection<IEncapsulateFieldCandidate> EncapsulateFieldCandidates { get; }

        public IReadOnlyCollection<IEncapsulateFieldAsUDTMemberCandidate> EncapsulateAsUserDefinedTypeMemberCandidates { get; }

        public IReadOnlyCollection<IObjectStateUDT> ObjectStateUDTCandidates { get; }

        private static List<IObjectStateUDT> LoadObjectStateUDTCandidates(IReadOnlyCollection<IEncapsulateFieldCandidate> fieldCandidates, IObjectStateUserDefinedTypeFactory factory, QualifiedModuleName qmn)
        {
            var objectStateUDTs = new List<IObjectStateUDT>();
            objectStateUDTs = fieldCandidates
                .OfType<IUserDefinedTypeCandidate>()
                .Where(fc => fc.Declaration.Accessibility == Accessibility.Private
                    && fc.Declaration.AsTypeDeclaration.Accessibility == Accessibility.Private)
                .Select(udtc => factory.Create(udtc))
                .ToList();

            //If more than one instance of a UserDefinedType is available, it is disqualified as 
            //a field to host the module's state.
            var multipleFieldsOfTheSameUDT = objectStateUDTs.ToLookup(os => os.Declaration.AsTypeDeclaration.IdentifierName);
            foreach (var duplicate in multipleFieldsOfTheSameUDT.Where(d => d.Count() > 1))
            {
                objectStateUDTs.RemoveAll(os => duplicate.Contains(os));
            }

            var defaultObjectStateUDT = factory.Create(qmn);
            objectStateUDTs.Add(defaultObjectStateUDT);

            return objectStateUDTs;
        }

        private static List<IEncapsulateFieldAsUDTMemberCandidate> LoadAsUDTMemberCandidates(IReadOnlyCollection<IEncapsulateFieldCandidate> fieldCandidates, IReadOnlyCollection<IObjectStateUDT> objectStateUDTCandidates)
        {
            var encapsulateAsUDTMembers = new List<IEncapsulateFieldAsUDTMemberCandidate>();
            var defaultObjectStateUDT = objectStateUDTCandidates.FirstOrDefault(os => !os.IsExistingDeclaration);

            foreach (var field in fieldCandidates)
            {
                var asUDT = new EncapsulateFieldAsUDTMemberCandidate(field, defaultObjectStateUDT);
                encapsulateAsUDTMembers.Add(asUDT);
            }

            return encapsulateAsUDTMembers;
        }
    }
}
