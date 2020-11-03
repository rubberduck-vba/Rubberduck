using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldCandidateSetsProviderFactory
    {
        IEncapsulateFieldCandidateSetsProvider Create(IDeclarationFinderProvider declarationFinderProvider,
            IEncapsulateFieldCandidateFactory encapsulateFieldCandidateFactory,
            QualifiedModuleName qualifiedModuleName);
    }

    public interface IEncapsulateFieldCandidateSetsProvider
    {
        IReadOnlyCollection<IEncapsulateFieldCandidate> EncapsulateFieldUseBackingFieldCandidates { get; }
        IReadOnlyCollection<IEncapsulateFieldAsUDTMemberCandidate> EncapsulateFieldUseBackingUDTMemberCandidates { get; }
        IReadOnlyCollection<IObjectStateUDT> ObjectStateFieldCandidates { get; }
    }

    /// <summary>
    /// EncapsulateFieldCandidateSetsProvider provides access to a sets of 
    /// EncapsulateField candidate instances to be shared among EncapsulateFieldRefactoringActions.
    /// </summary>
    public class EncapsulateFieldCandidateSetsProvider : IEncapsulateFieldCandidateSetsProvider
    {
        public EncapsulateFieldCandidateSetsProvider(IDeclarationFinderProvider declarationFinderProvider,
            IEncapsulateFieldCandidateFactory encapsulateFieldCandidateFactory,
            QualifiedModuleName qualifiedModuleName)
        {
            EncapsulateFieldUseBackingFieldCandidates = declarationFinderProvider.DeclarationFinder.Members(qualifiedModuleName, DeclarationType.Variable)
                .Where(v => v.ParentDeclaration is ModuleDeclaration
                    && !v.IsWithEvents)
                .Select(f => encapsulateFieldCandidateFactory.CreateFieldCandidate(f))
                .ToList();

            var objectStateUDTCandidates = EncapsulateFieldUseBackingFieldCandidates
                .OfType<IUserDefinedTypeCandidate>()
                .Where(fc => fc.Declaration.Accessibility == Accessibility.Private
                    && fc.Declaration.AsTypeDeclaration.Accessibility == Accessibility.Private)
                .Select(udtc => encapsulateFieldCandidateFactory.CreateObjectStateField(udtc))
                //If multiple fields of the same UserDefinedType exist, they are all disqualified as candidates to host a module's state.
                .ToLookup(objectStateUDTCandidate => objectStateUDTCandidate.Declaration.AsTypeDeclaration.IdentifierName)
                .Where(osc => osc.Count() == 1)
                .SelectMany(osc => osc)
                .ToList();

            var defaultObjectStateUDT = encapsulateFieldCandidateFactory.CreateDefaultObjectStateField(qualifiedModuleName);
            objectStateUDTCandidates.Add(defaultObjectStateUDT);
            ObjectStateFieldCandidates = objectStateUDTCandidates;

            EncapsulateFieldUseBackingUDTMemberCandidates = EncapsulateFieldUseBackingFieldCandidates
                .Select(fc => encapsulateFieldCandidateFactory.CreateUDTMemberCandidate(fc, defaultObjectStateUDT))
                .ToList();
        }

        public IReadOnlyCollection<IEncapsulateFieldCandidate> EncapsulateFieldUseBackingFieldCandidates { get; }

        public IReadOnlyCollection<IEncapsulateFieldAsUDTMemberCandidate> EncapsulateFieldUseBackingUDTMemberCandidates { get; }

        public IReadOnlyCollection<IObjectStateUDT> ObjectStateFieldCandidates { get; }
    }
}
