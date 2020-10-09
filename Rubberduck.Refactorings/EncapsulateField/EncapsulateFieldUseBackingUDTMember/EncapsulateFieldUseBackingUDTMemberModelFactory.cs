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
        /// <param name="objectStateField">Optional: <c>UserDefinedType</c> Field to host the Encapsulated Field(s)</param>
        EncapsulateFieldUseBackingUDTMemberModel Create(IEnumerable<FieldEncapsulationModel> fieldModels, Declaration objectStateField = null);

        /// <summary>
        /// Creates an <c>EncapsulateFieldUseBackingUDTMemberModel</c> based upon collection of
        /// <c>IEncapsulateFieldCandidate</c> instances.
        /// This function is intended for exclusive use by the <c>EncapsulateFieldModelFactory</c>
        /// </summary>
        /// <param name="objectStateField">Optional: <c>UserDefinedType</c> Field to host the Encapsulated Field(s)</param>
        EncapsulateFieldUseBackingUDTMemberModel Create(IEncapsulateFieldCandidateSetsProvider contextCollections, IEnumerable<FieldEncapsulationModel> fieldModels, Declaration objectStateField = null);
    }

    public class EncapsulateFieldUseBackingUDTMemberModelFactory : IEncapsulateFieldUseBackingUDTMemberModelFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IEncapsulateFieldCandidateFactory _candidatesFactory;
        private readonly IEncapsulateFieldCandidateSetsProviderFactory _candidateSetsFactory;
        private readonly IEncapsulateFieldConflictFinderFactory _conflictFinderFactory;

        public EncapsulateFieldUseBackingUDTMemberModelFactory(IDeclarationFinderProvider declarationFinderProvider,
            IEncapsulateFieldCandidateFactory candidatesFactory,
            IEncapsulateFieldCandidateSetsProviderFactory candidateSetsFactory,
            IEncapsulateFieldConflictFinderFactory encapsulateFieldConflictFinderFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _candidatesFactory = candidatesFactory;
            _candidateSetsFactory = candidateSetsFactory;
            _conflictFinderFactory = encapsulateFieldConflictFinderFactory;
        }

        public EncapsulateFieldUseBackingUDTMemberModel Create(IEnumerable<FieldEncapsulationModel> fieldModels, Declaration objectStateField)
        {
            if (!fieldModels.Any())
            {
                throw new ArgumentException();
            }

            var contextCollections = _candidateSetsFactory.Create(_declarationFinderProvider, _candidatesFactory, fieldModels.First().Declaration.QualifiedModuleName);

            return Create(contextCollections, fieldModels, objectStateField);
        }

        public EncapsulateFieldUseBackingUDTMemberModel Create(IEncapsulateFieldCandidateSetsProvider contextCollections, IEnumerable<FieldEncapsulationModel> fieldModels, Declaration objectStateField = null)
        {
            var fieldCandidates = contextCollections.EncapsulateFieldUseBackingUDTMemberCandidates.ToList();

            if (objectStateField != null
                && (objectStateField.Accessibility != Accessibility.Private
                    || !fieldCandidates.Any(c => c.Declaration == objectStateField && c.WrappedCandidate is IUserDefinedTypeCandidate)))
            {
                throw new ArgumentException("The object state Field must be a Private UserDefinedType");
            }

            var objectStateFieldCandidates = contextCollections.ObjectStateFieldCandidates;

            var defaultObjectStateUDT = objectStateFieldCandidates.FirstOrDefault(os => !os.IsExistingDeclaration);

            var targetStateUDT = DetermineObjectStateFieldTarget(defaultObjectStateUDT, objectStateField, objectStateFieldCandidates);

            foreach (var fieldModel in fieldModels)
            {
                var candidate = fieldCandidates.Single(c => c.Declaration.Equals(fieldModel.Declaration));
                candidate.EncapsulateFlag = true;
                candidate.IsReadOnly = fieldModel.IsReadOnly;
                if (fieldModel.PropertyIdentifier != null)
                {
                    candidate.PropertyIdentifier = fieldModel.PropertyIdentifier;
                }
            }

            var conflictsFinder = _conflictFinderFactory.Create(_declarationFinderProvider,
                contextCollections.EncapsulateFieldUseBackingFieldCandidates,
                contextCollections.ObjectStateFieldCandidates);

            fieldCandidates.ForEach(c => c.ConflictFinder = conflictsFinder);

            if (objectStateField == null && !targetStateUDT.IsExistingDeclaration)
            {
                conflictsFinder.AssignNoConflictIdentifiers(targetStateUDT);
            }

            fieldCandidates.ForEach(c => conflictsFinder.AssignNoConflictIdentifiers(c));

            return new EncapsulateFieldUseBackingUDTMemberModel(targetStateUDT, fieldCandidates, objectStateFieldCandidates)
            {
                ConflictFinder = conflictsFinder
            };
        }

        IObjectStateUDT DetermineObjectStateFieldTarget(IObjectStateUDT defaultObjectStateField, Declaration objectStateFieldTarget, IReadOnlyCollection<IObjectStateUDT> objectStateFieldCandidates)
        {
            var targetStateUDT = defaultObjectStateField;

            if (objectStateFieldTarget != null)
            {
                targetStateUDT = objectStateFieldCandidates.Single(osc => objectStateFieldTarget == osc.Declaration);
            }
            else
            {
                var preExistingDefaultUDTField = 
                    objectStateFieldCandidates.Where(osc => osc.TypeIdentifier == defaultObjectStateField.TypeIdentifier
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