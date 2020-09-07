using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
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
        EncapsulateFieldUseBackingUDTMemberModel Create(IEnumerable<EncapsulateFieldRequest> requests, Declaration userDefinedTypeTarget = null);

        /// <summary>
        /// Creates an <c>EncapsulateFieldUseBackingUDTMemberModel</c> based upon collection of
        /// <c>IEncapsulateFieldCandidate</c> instances created by <c>EncapsulateFieldCandidateCollectionFactory</c>.  
        /// This function is intended for exclusive use by <c>EncapsulateFieldModelFactory</c>
        /// </summary>
        EncapsulateFieldUseBackingUDTMemberModel Create(IReadOnlyCollection<IEncapsulateFieldCandidate> candidates, IEnumerable<EncapsulateFieldRequest> requests, Declaration userDefinedTypeTarget = null);
    }

    public class EncapsulateFieldUseBackingUDTMemberModelFactory : IEncapsulateFieldUseBackingUDTMemberModelFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IEncapsulateFieldCandidateCollectionFactory _fieldCandidateCollectionFactory;
        private readonly IObjectStateUserDefinedTypeFactory _objectStateUDTFactory;
        private readonly IEncapsulateFieldConflictFinderFactory _conflictFinderFactory;

        public EncapsulateFieldUseBackingUDTMemberModelFactory(IDeclarationFinderProvider declarationFinderProvider,
            IEncapsulateFieldCandidateCollectionFactory encapsulateFieldCandidateCollectionFactory,
            IObjectStateUserDefinedTypeFactory objectStateUserDefinedTypeFactory,
            IEncapsulateFieldConflictFinderFactory encapsulateFieldConflictFinderFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _fieldCandidateCollectionFactory = encapsulateFieldCandidateCollectionFactory;
            _objectStateUDTFactory = objectStateUserDefinedTypeFactory;
            _conflictFinderFactory = encapsulateFieldConflictFinderFactory;
        }

        public EncapsulateFieldUseBackingUDTMemberModel Create(IEnumerable<EncapsulateFieldRequest> requests, Declaration clientTarget)
        {
            if (!requests.Any())
            {
                throw new ArgumentException();
            }
            var fieldCandidates = _fieldCandidateCollectionFactory.Create(requests.First().Declaration.QualifiedModuleName);
            return Create(fieldCandidates, requests, clientTarget);
        }

        public EncapsulateFieldUseBackingUDTMemberModel Create(IReadOnlyCollection<IEncapsulateFieldCandidate> candidates, IEnumerable<EncapsulateFieldRequest> requests, Declaration clientTarget = null)
        {
            if (clientTarget != null
                && (clientTarget.Accessibility != Accessibility.Private
                    || !candidates.Any(c => c.Declaration == clientTarget && c is IUserDefinedTypeCandidate)))
            {
                throw new ArgumentException("The object state Field must be a Private UserDefinedType");
            }

            var fieldCandidates = candidates.ToList();

            var objectStateUDTs = fieldCandidates
                .Where(c => c is IUserDefinedTypeCandidate udt && udt.IsObjectStateUDTCandidate)
                .Select(udtc => _objectStateUDTFactory.Create(udtc as IUserDefinedTypeCandidate))
                .ToList();

            var defaultObjectStateUDT = _objectStateUDTFactory.Create(fieldCandidates.First().QualifiedModuleName);

            objectStateUDTs.Add(defaultObjectStateUDT);

            var targetStateUDT = DetermineObjectStateUDTTarget(defaultObjectStateUDT, clientTarget, objectStateUDTs);

            foreach (var request in requests)
            {
                var candidate = fieldCandidates.Single(c => c.Declaration.Equals(request.Declaration));
                request.ApplyRequest(candidate);
            }

            var conflictsFinder = _conflictFinderFactory.CreateEncapsulateFieldUseBackingUDTMemberConflictFinder(candidates, objectStateUDTs)
                as IEncapsulateFieldUseBackingUDTMemberConflictFinder;

            fieldCandidates.ForEach(c => c.ConflictFinder = conflictsFinder);

            if (clientTarget == null && !targetStateUDT.IsExistingDeclaration)
            {
                conflictsFinder.AssignNoConflictIdentifiers(targetStateUDT, _declarationFinderProvider);
            }

            var udtMemberCandidates =
                fieldCandidates.Select(c => new EncapsulateFieldAsUDTMemberCandidate(c, targetStateUDT)).ToList();

            udtMemberCandidates.ForEach(c => conflictsFinder.AssignNoConflictIdentifiers(c));

            return new EncapsulateFieldUseBackingUDTMemberModel(targetStateUDT, udtMemberCandidates, objectStateUDTs)
            {
                ConflictFinder = conflictsFinder
            };
        }

        IObjectStateUDT DetermineObjectStateUDTTarget(IObjectStateUDT defaultObjectStateUDT, Declaration clientTarget, List<IObjectStateUDT> objectStateUDTs)
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
