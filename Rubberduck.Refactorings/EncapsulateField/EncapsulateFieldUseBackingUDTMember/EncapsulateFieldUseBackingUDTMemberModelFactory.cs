using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings
{
    public interface IEncapsulateFieldUseBackingUDTMemberModelFactory : IEncapsulateFieldModelsFactory<EncapsulateFieldUseBackingUDTMemberModel>
    { }

    public class EncapsulateFieldUseBackingUDTMemberModelFactory : IEncapsulateFieldUseBackingUDTMemberModelFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IEncapsulateFieldCandidateFactory _fieldCandidateFactory;

        public EncapsulateFieldUseBackingUDTMemberModelFactory(IDeclarationFinderProvider declarationFinderProvider, IEncapsulateFieldCandidateFactory fieldCandidateFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _fieldCandidateFactory = fieldCandidateFactory;
        }

        public EncapsulateFieldUseBackingUDTMemberModel Create(QualifiedModuleName qmn)        {
            var fields = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(v => v.ParentDeclaration is ModuleDeclaration
                    && !v.IsWithEvents);

            var candidates = fields.Select(f => _fieldCandidateFactory.Create(f));

            var objectStateUDTCandidates = candidates.Where(c => c is IUserDefinedTypeCandidate udt
                && udt.CanBeObjectStateUDT)
                .Select(udtc => new ObjectStateUDT(udtc as IUserDefinedTypeCandidate));

            return Create(candidates, objectStateUDTCandidates);
        }

        public EncapsulateFieldUseBackingUDTMemberModel Create(IEnumerable<IEncapsulateFieldCandidate> candidates, IEnumerable<IObjectStateUDT> objectStateUDTCandidates)
        {
            var fieldCandidates = new List<IEncapsulateFieldCandidate>(candidates);
            var objectStateFieldCandidates = new List<IObjectStateUDT>(objectStateUDTCandidates);

            var udtMemberCandidates = new List<IUserDefinedTypeMemberCandidate>();
            fieldCandidates.ForEach(c => LoadUDTMembers(udtMemberCandidates, c));

            var conflictsFinder = new ConvertFieldsToUDTMembersStrategyConflictFinder(_declarationFinderProvider, candidates, udtMemberCandidates, objectStateUDTCandidates);
            fieldCandidates.ForEach(c => c.ConflictFinder = conflictsFinder);

            var defaultObjectStateUDT = CreateStateUDTField(candidates.First().QualifiedModuleName);
            conflictsFinder.AssignNoConflictIdentifiers(defaultObjectStateUDT, _declarationFinderProvider);

            var convertedToUDTMemberCandidates = new List<IConvertToUDTMember>();
            foreach (var field in candidates)
            {
                if (field is ConvertToUDTMember cm)
                {
                    convertedToUDTMemberCandidates.Add(cm);
                    continue;
                }
                convertedToUDTMemberCandidates.Add(new ConvertToUDTMember(field, defaultObjectStateUDT));
            }

            return new EncapsulateFieldUseBackingUDTMemberModel(convertedToUDTMemberCandidates, defaultObjectStateUDT, objectStateUDTCandidates, _declarationFinderProvider)
            {
                ConflictFinder = conflictsFinder
            };
        }

        private IObjectStateUDT CreateStateUDTField(QualifiedModuleName qualifiedModuleName)
        {
            var stateUDT = new ObjectStateUDT(qualifiedModuleName) as IObjectStateUDT;
            stateUDT.IsSelected = true;

            return stateUDT;
        }

        private void ResolveConflict(IEncapsulateFieldConflictFinder conflictFinder, IEncapsulateFieldCandidate candidate)
        {
            conflictFinder.AssignNoConflictIdentifiers(candidate);
            if (candidate is IUserDefinedTypeCandidate udtCandidate)
            {
                foreach (var member in udtCandidate.Members)
                {
                    conflictFinder.AssignNoConflictIdentifiers(member);
                    if (member.WrappedCandidate is IUserDefinedTypeCandidate childUDT
                        && childUDT.Declaration.AsTypeDeclaration.HasPrivateAccessibility())
                    {
                        ResolveConflict(conflictFinder, childUDT);
                    }
                }
            }
        }

        private void LoadUDTMembers(List<IUserDefinedTypeMemberCandidate> udtMembers, IEncapsulateFieldCandidate candidate)
        {
            if (candidate is IUserDefinedTypeCandidate udtCandidate)
            {
                foreach (var member in udtCandidate.Members)
                {
                    udtMembers.Add(member);
                    if (member.WrappedCandidate is IUserDefinedTypeCandidate childUDT
                        && childUDT.Declaration.AsTypeDeclaration.HasPrivateAccessibility())
                    {
                        LoadUDTMembers(udtMembers, childUDT);
                    }
                }
            }
        }
    }
}
