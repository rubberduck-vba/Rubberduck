using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingField;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings
{
    public interface IEncapsulateFieldUseBackingFieldModelFactory : IEncapsulateFieldModelsFactory<EncapsulateFieldUseBackingFieldModel>
    { }

    public class EncapsulateFieldUseBackingFieldModelFactory : IEncapsulateFieldUseBackingFieldModelFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IEncapsulateFieldCandidateFactory _fieldCandidateFactory;

        public EncapsulateFieldUseBackingFieldModelFactory(IDeclarationFinderProvider declarationFinderProvider, 
            IEncapsulateFieldCandidateFactory fieldCandidateFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _fieldCandidateFactory = fieldCandidateFactory;
        }

        public EncapsulateFieldUseBackingFieldModel Create(QualifiedModuleName qmn)
        {
            var fields = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(v => v.ParentDeclaration is ModuleDeclaration
                    && !v.IsWithEvents);

            var candidates = fields.Select(f => _fieldCandidateFactory.Create(f));

            var objectStateUDTCandidates = candidates.Where(c => c is IUserDefinedTypeCandidate udt && udt.CanBeObjectStateUDT)
                .Select(udtc => new ObjectStateUDT(udtc as IUserDefinedTypeCandidate));

            return Create(candidates, objectStateUDTCandidates);
        }

        public EncapsulateFieldUseBackingFieldModel Create(IEnumerable<IEncapsulateFieldCandidate> candidates, IEnumerable<IObjectStateUDT> objectStateUDTCandidates)
        {
            var fieldCandidates = new List<IEncapsulateFieldCandidate>(candidates);
            var objectStateFieldCandidates = new List<IObjectStateUDT>(objectStateUDTCandidates);
            var udtMemberCandidates = new List<IUserDefinedTypeMemberCandidate>();

            fieldCandidates.ForEach(c => LoadUDTMembers(udtMemberCandidates, c));

            var conflictsFinder = new UseBackingFieldsStrategyConflictFinder(_declarationFinderProvider, candidates, udtMemberCandidates);
            fieldCandidates.ForEach(c => c.ConflictFinder = conflictsFinder);

            return new EncapsulateFieldUseBackingFieldModel(candidates, _declarationFinderProvider)
            {
                ConflictFinder = conflictsFinder
            };
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
