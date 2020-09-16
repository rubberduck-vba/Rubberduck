using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings
{
    public interface IEncapsulateFieldUseBackingUDTMemberModelFactory
    {
        /// <summary>
        /// Creates an <c>EncapsulateFieldUseBackingUDTMemberModel</c> used by the <c>EncapsulateFieldUseBackingUDTMemberRefactoringAction</c>.
        /// </summary>
        /// <param name="clientTarget">Optional: <c>UserDefinedType</c> Field to include the Encapsulated Field(s)</param>
        EncapsulateFieldUseBackingUDTMemberModel Create(IEnumerable<FieldEncapsulationModel> requests, Declaration userDefinedTypeTarget = null);

        /// <summary>
        /// Creates an <c>EncapsulateFieldUseBackingUDTMemberModel</c> based upon collection of
        /// <c>IEncapsulateFieldCandidate</c> instances created by <c>EncapsulateFieldCandidateCollectionFactory</c>.  
        /// This function is intended for exclusive use by <c>EncapsulateFieldModelFactory</c>
        /// </summary>
        EncapsulateFieldUseBackingUDTMemberModel Create(IEncapsulateFieldCollectionsProvider collectionsProvider, IEnumerable<FieldEncapsulationModel> requests, Declaration userDefinedTypeTarget = null);
    }

    public class EncapsulateFieldUseBackingUDTMemberModelFactory : IEncapsulateFieldUseBackingUDTMemberModelFactory
    {
        private readonly IEncapsulateFieldCollectionsProviderFactory _encapsulateFieldCollectionsProviderFactory;
        private readonly IEncapsulateFieldConflictFinderFactory _conflictFinderFactory;

        public EncapsulateFieldUseBackingUDTMemberModelFactory(
            IEncapsulateFieldCollectionsProviderFactory encapsulateFieldCollectionsProviderFactory,
            IEncapsulateFieldConflictFinderFactory encapsulateFieldConflictFinderFactory)
        {
            _conflictFinderFactory = encapsulateFieldConflictFinderFactory;
            _encapsulateFieldCollectionsProviderFactory = encapsulateFieldCollectionsProviderFactory;
        }

        public EncapsulateFieldUseBackingUDTMemberModel Create(IEnumerable<FieldEncapsulationModel> requests, Declaration clientTarget)
        {
            if (!requests.Any())
            {
                throw new ArgumentException();
            }

            var collectionsProvider = _encapsulateFieldCollectionsProviderFactory.Create(requests.First().Declaration.QualifiedModuleName);

            return Create(collectionsProvider, requests, clientTarget);
        }

        public EncapsulateFieldUseBackingUDTMemberModel Create(IEncapsulateFieldCollectionsProvider collectionsProvider, IEnumerable<FieldEncapsulationModel> requests, Declaration clientTarget = null)
        {
            var asUDTMemberCandidates = collectionsProvider.EncapsulateAsUserDefinedTypeMemberCandidates.ToList();

            if (clientTarget != null
                && (clientTarget.Accessibility != Accessibility.Private
                    || !asUDTMemberCandidates.Any(c => c.Declaration == clientTarget && c.WrappedCandidate is IUserDefinedTypeCandidate)))
            {
                throw new ArgumentException("The object state Field must be a Private UserDefinedType");
            }

            var objectStateUDTs = collectionsProvider.ObjectStateUDTCandidates;

            var defaultObjectStateUDT = objectStateUDTs.FirstOrDefault(os => !os.IsExistingDeclaration);

            var targetStateUDT = DetermineObjectStateUDTTarget(defaultObjectStateUDT, clientTarget, objectStateUDTs);

            foreach (var request in requests)
            {
                var candidate = asUDTMemberCandidates.Single(c => c.Declaration.Equals(request.Declaration));
                candidate.EncapsulateFlag = true;
                candidate.IsReadOnly = request.IsReadOnly;
                if (request.PropertyIdentifier != null)
                {
                    candidate.PropertyIdentifier = request.PropertyIdentifier;
                }
            }

            var conflictsFinder = _conflictFinderFactory.Create(collectionsProvider);

            asUDTMemberCandidates.ForEach(c => c.ConflictFinder = conflictsFinder);

            if (clientTarget == null && !targetStateUDT.IsExistingDeclaration)
            {
                conflictsFinder.AssignNoConflictIdentifiers(targetStateUDT);
            }

            asUDTMemberCandidates.ForEach(c => conflictsFinder.AssignNoConflictIdentifiers(c));

            return new EncapsulateFieldUseBackingUDTMemberModel(targetStateUDT, asUDTMemberCandidates, objectStateUDTs)
            {
                ConflictFinder = conflictsFinder
            };
        }

        IObjectStateUDT DetermineObjectStateUDTTarget(IObjectStateUDT defaultObjectStateUDT, Declaration clientTarget, IReadOnlyCollection<IObjectStateUDT> objectStateUDTs)
        {
            var targetStateUDT = defaultObjectStateUDT;

            if (clientTarget != null)
            {
                targetStateUDT = objectStateUDTs.Single(osc => clientTarget == osc.Declaration);
            }
            else
            {
                var preExistingDefaultUDTField = 
                    objectStateUDTs.Where(osc => osc.TypeIdentifier == defaultObjectStateUDT.TypeIdentifier
                        && osc.IsExistingDeclaration);

                if (preExistingDefaultUDTField.Any() && preExistingDefaultUDTField.Count() == 1)
                {
                    targetStateUDT = preExistingDefaultUDTField.First();
                }
            }

            targetStateUDT.IsSelected = true;

            return targetStateUDT;
        }
    }
}
