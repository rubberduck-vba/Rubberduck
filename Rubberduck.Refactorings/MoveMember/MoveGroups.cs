using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveMember.Extensions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMoveMemberGroups
    {
        IEnumerable<Declaration> CallChainDeclarations(IEnumerable<Declaration> selected);
        IEnumerable<Declaration> AllParticipants { get; }
        IEnumerable<Declaration> Selected { get; }
        IEnumerable<Declaration> SupportFields { get; }
        IEnumerable<Declaration> SupportConstants { get; }
        IEnumerable<Declaration> SupportLetSetGetPropertyGroups { get; }
        IEnumerable<Declaration> SupportMethods { get; }
        IEnumerable<Declaration> ExclusiveSupportFields { get; }
        IEnumerable<Declaration> ExclusiveSupportConstants { get; }
        IEnumerable<Declaration> ExclusiveSupportLetSetGetPropertyGroups { get; }
        IEnumerable<Declaration> ExclusiveSupportMethods { get; }
        IEnumerable<Declaration> NonExclusiveSupportFields { get; }
        IEnumerable<Declaration> NonExclusiveSupportConstants { get; }
        IEnumerable<Declaration> NonExclusiveSupportLetSetGetPropertyGroups { get; }
        IEnumerable<Declaration> NonExclusiveSupportMethods { get; }
        IEnumerable<Declaration> AllExclusiveSupportDeclarations { get; }
        IEnumerable<Declaration> AllNonExclusiveSupportDeclarations { get; }
        IEnumerable<Declaration> SourceModuleDeclarations { get; }
    }

    public class MoveMemberGroups : IMoveMemberGroups
    {
        private List<Declaration> _allSourceMembers;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public MoveMemberGroups(IEnumerable<Declaration> selectedDeclarations, IDeclarationFinderProvider declarationFinderProvider)
        {
            Debug.Assert(selectedDeclarations.Count() > 0);
            _declarationFinderProvider = declarationFinderProvider;

            Selected = selectedDeclarations.ToList();
            AllParticipants = CallChainDeclarations(selectedDeclarations).ToList();

            _allSourceMembers = declarationFinderProvider.DeclarationFinder.Members(Selected.First().QualifiedModuleName).ToList();

            var exclusiveSupportProperties = new List<Declaration>();
            var sameNameLetSetGetPropertyGroups = SupportLetSetGetPropertyGroups.GroupBy(key => key.IdentifierName);
            foreach (var lsgPropertyGroup in sameNameLetSetGetPropertyGroups)
            {
                if (lsgPropertyGroup.AllReferences().All(seRefs => AllParticipants.Contains(seRefs.ParentScoping)))
                {
                    exclusiveSupportProperties.AddRange(lsgPropertyGroup);
                }
            }

            ExclusiveSupportLetSetGetPropertyGroups = exclusiveSupportProperties;
        }

        public IEnumerable<Declaration> CallChainDeclarations(IEnumerable<Declaration> selected)
        {
            var allMembers = _declarationFinderProvider.DeclarationFinder.Members(selected.First().QualifiedModuleName);
            var participatingDeclarations = new HashSet<Declaration>();

            foreach (var element in selected)
            {
                participatingDeclarations.Add(element);
            }

            var allReferences = allMembers.AllReferences().ToList();

            var guard = 0;
            var newElements = selected;
            while (newElements.Any() && guard++ < 100)
            {
                newElements = allReferences
                        .Where(rf => newElements.Contains(rf.ParentScoping)
                                && rf.ParentScoping != rf.Declaration)
                        .Select(rf => rf.Declaration)
                .ToList();

                foreach (var element in newElements)
                {
                    if (IsMoveableDeclaration(element))
                    {
                        participatingDeclarations.Add(element);
                    }
                }
            }
            Debug.Assert(guard < 100);
            if (guard >= 100)
            {
                throw new MoveMemberUnsupportedMoveException(selected.FirstOrDefault());
            }

            var supportAsTypes = participatingDeclarations.Where(apd => apd.AsTypeDeclaration != null)
                .Select(atd => atd.AsTypeDeclaration.DeclarationType);

            if (supportAsTypes.Any())
            {
                var supports = allMembers.Where(m => supportAsTypes.Contains(m.DeclarationType));
                foreach (var support in supports)
                {
                    participatingDeclarations.Add(support);
                }
            }
            participatingDeclarations.RemoveWhere(m => m.DeclarationType == DeclarationType.Module);

            return participatingDeclarations;
        }

        public IEnumerable<Declaration> Selected { get; }

        public IEnumerable<Declaration> AllParticipants { get; }

        public IEnumerable<Declaration> SourceModuleDeclarations => _allSourceMembers;

        public IEnumerable<Declaration> SupportConstants
            => AllParticipants.Where(p => p.IsModuleConstant() && !Selected.Contains(p));

        public IEnumerable<Declaration> SupportFields
            => AllParticipants.Where(p => p.IsField() && !Selected.Contains(p));

        public IEnumerable<Declaration> SupportLetSetGetPropertyGroups
            => AllParticipants.Where(p => p.IsMember()
                    && p.DeclarationType.HasFlag(DeclarationType.Property)
                    && !Selected.Contains(p))
                .SelectMany(prop => _allSourceMembers.Where(p => p.DeclarationType.HasFlag(DeclarationType.Property) && p.IdentifierName == prop.IdentifierName));

        public IEnumerable<Declaration> SupportMethods
            => AllParticipants.Where(p => p.IsMember()
                    && !p.DeclarationType.HasFlag(DeclarationType.Property)
                    && !Selected.Contains(p));

        public IEnumerable<Declaration> ExclusiveSupportFields => SupportFields
                .Where(se => se.References.All(rf => AllParticipants.Contains(rf.ParentScoping)));

        public IEnumerable<Declaration> NonExclusiveSupportFields
            => SupportFields.Except(ExclusiveSupportFields);

        public IEnumerable<Declaration> ExclusiveSupportConstants => SupportConstants
                .Where(se => se.References.All(rf => AllParticipants.Contains(rf.ParentScoping)));

        public IEnumerable<Declaration> NonExclusiveSupportConstants
            => SupportConstants.Except(ExclusiveSupportConstants);

        public IEnumerable<Declaration> ExclusiveSupportMethods => SupportMethods
                .Where(se => se.References.All(seRefs => AllParticipants.Contains(seRefs.ParentScoping)));

        public IEnumerable<Declaration> NonExclusiveSupportMethods
            => SupportMethods.Except(ExclusiveSupportMethods);

        public IEnumerable<Declaration> ExclusiveSupportLetSetGetPropertyGroups { get; }

        public IEnumerable<Declaration> NonExclusiveSupportLetSetGetPropertyGroups
            => SupportLetSetGetPropertyGroups.Except(ExclusiveSupportLetSetGetPropertyGroups);

        public IEnumerable<Declaration> AllExclusiveSupportDeclarations
            => ExclusiveSupportFields
                .Concat(ExclusiveSupportConstants)
                .Concat(ExclusiveSupportMethods)
                .Concat(ExclusiveSupportLetSetGetPropertyGroups);

        public IEnumerable<Declaration> AllNonExclusiveSupportDeclarations
            => NonExclusiveSupportMethods
                .Concat(NonExclusiveSupportLetSetGetPropertyGroups)
                .Concat(NonExclusiveSupportFields)
                .Concat(NonExclusiveSupportConstants);

        private static bool IsMoveableDeclaration(Declaration declaration)
        {
            return (declaration.DeclarationType.HasFlag(DeclarationType.Member)
                        || declaration.IsField()
                        || declaration.IsModuleConstant());
        }
    }
}
