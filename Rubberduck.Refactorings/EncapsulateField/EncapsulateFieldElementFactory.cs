using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldElementFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IEncapsulateFieldNamesValidator _validator;
        private QualifiedModuleName _targetQMN;

        public EncapsulateFieldElementFactory(IDeclarationFinderProvider declarationFinderProvider, QualifiedModuleName targetQMN, IEncapsulateFieldNamesValidator validator)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _validator = validator;
            _targetQMN = targetQMN;
        }

        public IStateUDT CreateStateUDTField()
        {
            var stateUDT = new StateUDT(_targetQMN, _validator) as IStateUDT;

            stateUDT = SetNonConflictIdentifier(stateUDT, c => { return _validator.IsConflictingStateUDTFieldIdentifier(stateUDT); }, (s) => { stateUDT.FieldIdentifier = s; }, () => stateUDT.FieldIdentifier, _validator);

            stateUDT = SetNonConflictIdentifier(stateUDT, c => { return _validator.IsConflictingStateUDTTypeIdentifier(stateUDT); }, (s) => { stateUDT.TypeIdentifier = s; }, () => stateUDT.TypeIdentifier, _validator);

            return stateUDT;
        }

        public IEncapsulateFieldCandidate CreateEncapsulationCandidate(Declaration target)
        {
            Debug.Assert(!target.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember));

            IEncapsulateFieldCandidate candidate = CreateCandidate(target);

            candidate = ApplyTypeSpecificAttributes(candidate);
            if (candidate is IUserDefinedTypeCandidate udtVariable)
            {
                (Declaration udtDeclaration, IEnumerable<Declaration> udtMembers) = GetUDTAndMembersForField(udtVariable);

                udtVariable.TypeDeclarationIsPrivate = udtDeclaration.HasPrivateAccessibility();

                foreach (var udtMemberDeclaration in udtMembers)
                {
                    var candidateUDTMember = new UserDefinedTypeMemberCandidate(CreateCandidate(udtMemberDeclaration), udtVariable, _validator) as IUserDefinedTypeMemberCandidate;

                    candidateUDTMember = ApplyTypeSpecificAttributes(candidateUDTMember);

                    udtVariable.AddMember(candidateUDTMember);
                }
            }

            _validator.RegisterFieldCandidate(candidate);

            //candidate = _validator.AssignNoConflictIdentifier(candidate, DeclarationType.Property);
            //candidate = _validator.AssignNoConflictIdentifier(candidate, DeclarationType.Variable);

            return candidate;
        }

        private IEncapsulateFieldCandidate CreateCandidate(Declaration target)
        {
            if (target.IsUserDefinedTypeField())
            {
                return new UserDefinedTypeCandidate(target, _validator);
            }
            else if (target.IsArray)
            {
                return new ArrayCandidate(target, _validator);
            }
            return new EncapsulateFieldCandidate(target, _validator);
        }

        public IEnumerable<IEncapsulateFieldCandidate> CreateEncapsulationCandidates()
        {
            var fieldDeclarations = _declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN)
                .Where(v => v.IsMemberVariable() && !v.IsWithEvents);

            var candidates = new List<IEncapsulateFieldCandidate>();
            foreach (var field in fieldDeclarations)
            {
                var fieldEncapsulationCandidate = CreateEncapsulationCandidate(field);

                candidates.Add(fieldEncapsulationCandidate);
            }

            return candidates;
        }

        private IStateUDT SetNonConflictIdentifier(IStateUDT candidate, Predicate<IStateUDT> conflictDetector, Action<string> setValue, Func<string> getIdentifier, IEncapsulateFieldNamesValidator validator)
        {
            var isConflictingIdentifier = conflictDetector(candidate);
            for (var count = 1; count < 10 && isConflictingIdentifier; count++)
            {
                setValue(getIdentifier().IncrementEncapsulationIdentifier());
                isConflictingIdentifier = conflictDetector(candidate);
            }
            return candidate;
        }

        private T ApplyTypeSpecificAttributes<T>(T candidate) where T: IEncapsulateFieldCandidate
        {
            /*
             * Default values are:
             * candidate.ImplementLetSetterType = true;
             * candidate.ImplementSetSetterType = false;
             * candidate.CanBeReadWrite = true;
             * candidate.IsReadOnly = false;
             */

            if (candidate.Declaration.IsEnumField())
            {
                //5.3.1 The declared type of a function declaration may not be a private enum name.
                if (candidate.Declaration.AsTypeDeclaration.HasPrivateAccessibility())
                {
                    candidate.AsTypeName_Property = Tokens.Long;
                }
            }
            else if (candidate.Declaration.AsTypeName.Equals(Tokens.Variant)
                && !candidate.Declaration.IsArray)
            {
                candidate.ImplementLet = true;
                candidate.ImplementSet = true;
            }
            else if (candidate.Declaration.IsObject)
            {
                candidate.ImplementLet = false;
                candidate.ImplementSet = true;
            }
            return candidate;
        }

        private (Declaration TypeDeclaration, IEnumerable<Declaration> Members) GetUDTAndMembersForField(IUserDefinedTypeCandidate udtField)
        {
            var userDefinedTypeDeclaration = _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.UserDefinedType)
                .Where(ut => ut.IdentifierName.Equals(udtField.AsTypeName_Field)
                    && ut.QualifiedModuleName == udtField.QualifiedModuleName)
                .SingleOrDefault();

            if (userDefinedTypeDeclaration is null)
            {
                userDefinedTypeDeclaration = _declarationFinderProvider.DeclarationFinder
                    .UserDeclarations(DeclarationType.UserDefinedType)
                    .Where(ut => ut.IdentifierName.Equals(udtField.AsTypeName_Field)
                        && ut.ProjectId == udtField.Declaration.ProjectId
                        && ut.Accessibility != Accessibility.Private)
                .SingleOrDefault();
            }

            var udtMembers = _declarationFinderProvider.DeclarationFinder
               .UserDeclarations(DeclarationType.UserDefinedTypeMember)
               .Where(utm => userDefinedTypeDeclaration == utm.ParentDeclaration);

            return (userDefinedTypeDeclaration, udtMembers);
        }
    }
}
