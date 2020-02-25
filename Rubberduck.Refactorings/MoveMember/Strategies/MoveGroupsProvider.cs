using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMoveGroupsProvider
    {
        IReadOnlyCollection<Declaration> AllParticipants { get; }
        IReadOnlyCollection<Declaration> CallTreeRoots { get; }
        IReadOnlyCollection<Declaration> ExclusiveSupportDeclarations { get; }
        IReadOnlyCollection<Declaration> NonExclusiveSupportDeclarations { get; }
        IReadOnlyCollection<Declaration> SupportMembers { get; }
    }

    public class MoveGroupsProvider : IMoveGroupsProvider
    {
        private List<Declaration> _allModuleDeclarations;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        private Dictionary<string, List<Declaration>> _exclusiveSupportLetSetGetProperties;
        private Dictionary<string, List<Declaration>> _nonExclusiveSupportLetSetGetProperties;

        public MoveGroupsProvider(IEnumerable<Declaration> selectedDeclarations, IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;

            _callTreeRoots = new List<Declaration>();
            _allModuleDeclarations = new List<Declaration>();
            _allParticipants = new List<Declaration>();
            _exclusiveSupportLetSetGetProperties = new Dictionary<string, List<Declaration>>();
            _nonExclusiveSupportLetSetGetProperties = new Dictionary<string, List<Declaration>>();

            if (!selectedDeclarations.Any()) { return; }

            _callTreeRoots = selectedDeclarations.ToList();
            _allModuleDeclarations = declarationFinderProvider.DeclarationFinder.Members(selectedDeclarations.First().QualifiedModuleName).ToList();

            //Modify _callTreeRoots to remove selected declarations that
            //are a support member of any of the other selected declarations
            foreach (var selected in selectedDeclarations)
            {
                var allSupportParticipants = CallTreeDeclarations(new Declaration[] { selected }, true).ToList();
                _callTreeRoots.RemoveAll(d => allSupportParticipants.Contains(d));
            }

            _allParticipants = CallTreeDeclarations(_callTreeRoots).ToList();

            var letSetGetPropertyGroups = _allModuleDeclarations
                .Where(p => p.DeclarationType.HasFlag(DeclarationType.Property))
                .GroupBy(key => key.IdentifierName);

            _exclusiveSupportLetSetGetProperties = new Dictionary<string, List<Declaration>>();
            _nonExclusiveSupportLetSetGetProperties = new Dictionary<string, List<Declaration>>();

            foreach (var lsgPropertyGroup in letSetGetPropertyGroups)
            {
                if (_allParticipants.ContainsParentScopesForAllReferences(lsgPropertyGroup.AllReferences()))
                {
                    var propertyGroup = lsgPropertyGroup.Except(_callTreeRoots);
                    _exclusiveSupportLetSetGetProperties.Add(lsgPropertyGroup.Key, propertyGroup.ToList());
                }
                else if (_allParticipants.ContainsParentScopeForAnyReference(lsgPropertyGroup.AllReferences()))
                {
                    var propertyGroup = lsgPropertyGroup.Except(_callTreeRoots);
                    _nonExclusiveSupportLetSetGetProperties.Add(lsgPropertyGroup.Key, propertyGroup.ToList());
                }
            }
        }

        private IReadOnlyCollection<Declaration> CallTreeDeclarations(IEnumerable<Declaration> definingDeclarations, bool supportDeclarationsOnly = false)
        {
            if (!definingDeclarations.Any())
            {
                return new HashSet<Declaration>();
            }

            var allElements = _declarationFinderProvider.DeclarationFinder.Members(definingDeclarations.First().QualifiedModuleName);
            var participatingDeclarations = new HashSet<Declaration>();

            foreach (var element in definingDeclarations)
            {
                participatingDeclarations.Add(element);
            }

            var allReferences = allElements.AllReferences().ToList();

            var maxIterations = 100;
            var guard = 0;
            var newElements = definingDeclarations;
            while (newElements.Any() && guard++ < maxIterations)
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

            Debug.Assert(guard < maxIterations);
            if (guard >= maxIterations)
            {
                throw new MoveMemberUnsupportedMoveException(definingDeclarations.FirstOrDefault());
            }

            var supportAsTypes = participatingDeclarations.Where(apd => apd.AsTypeDeclaration != null)
                .Select(atd => atd.AsTypeDeclaration.DeclarationType);

            if (supportAsTypes.Any())
            {
                var supports = allElements.Where(m => supportAsTypes.Contains(m.DeclarationType));
                foreach (var support in supports)
                {
                    participatingDeclarations.Add(support);
                }
            }
            participatingDeclarations.RemoveWhere(m => m.DeclarationType == DeclarationType.Module);

            return supportDeclarationsOnly
                ? participatingDeclarations.Except(definingDeclarations).ToList()
                : participatingDeclarations.ToList();
        }

        private List<Declaration> _callTreeRoots;
        public IReadOnlyCollection<Declaration> CallTreeRoots => _callTreeRoots;

        private List<Declaration> _allParticipants;
        public IReadOnlyCollection<Declaration> AllParticipants => _allParticipants;

        public IReadOnlyCollection<Declaration> ExclusiveSupportDeclarations
            => ExclusiveSupportFields
                .Concat(ExclusiveSupportConstants)
                .Concat(ExclusiveSupportMethods)
                .Concat(ExclusiveSupportProperties).ToList();

        public IReadOnlyCollection<Declaration> NonExclusiveSupportDeclarations
            => NonExclusiveSupportMethods
                .Concat(NonExclusiveSupportProperties)
                .Concat(NonExclusiveSupportFields)
                .Concat(NonExclusiveSupportConstants).ToList();

        public IReadOnlyCollection<Declaration> SupportMembers
            => AllParticipants.Where(p => p.IsMember()
                    && !_callTreeRoots.Contains(p)).ToList();

        private IReadOnlyCollection<Declaration> ExclusiveSupportFields => SupportFields
                .Where(se => se.References.All(rf => AllParticipants.Contains(rf.ParentScoping))).ToList();

        private IReadOnlyCollection<Declaration> NonExclusiveSupportFields
            => SupportFields.Except(ExclusiveSupportFields).ToList();

        private IReadOnlyCollection<Declaration> ExclusiveSupportConstants => SupportConstants
                .Where(se => se.References.All(rf => AllParticipants.Contains(rf.ParentScoping))).ToList();

        private IReadOnlyCollection<Declaration> NonExclusiveSupportConstants
            => SupportConstants.Except(ExclusiveSupportConstants).ToList();

        private IReadOnlyCollection<Declaration> ExclusiveSupportMethods => SupportNonPropertyMembers
                .Where(se => se.References.All(seRefs => AllParticipants.Contains(seRefs.ParentScoping))).ToList();

        private IReadOnlyCollection<Declaration> NonExclusiveSupportMethods
            => SupportNonPropertyMembers.Except(ExclusiveSupportMethods).ToList();

        private IReadOnlyCollection<Declaration> ExclusiveSupportProperties
            => _exclusiveSupportLetSetGetProperties.Values.SelectMany(v => v).ToList();

        private IReadOnlyCollection<Declaration> NonExclusiveSupportProperties
            => _nonExclusiveSupportLetSetGetProperties.Values.SelectMany(v => v).ToList();

        private IEnumerable<Declaration> SupportNonPropertyMembers
            => AllParticipants.Where(p => p.IsMember()
                    && !p.DeclarationType.HasFlag(DeclarationType.Property)
                    && !_callTreeRoots.Contains(p));

        private IEnumerable<Declaration> SupportConstants
            => AllParticipants.Where(p => p.IsModuleConstant() && !_callTreeRoots.Contains(p));

        private IEnumerable<Declaration> SupportFields
            => AllParticipants.Where(p => p.IsField() && !_callTreeRoots.Contains(p));

        private IEnumerable<Declaration> SupportProperties
            => AllParticipants.Where(p => p.DeclarationType.HasFlag(DeclarationType.Property)
                    && !_callTreeRoots.Contains(p))
                .SelectMany(prop => _allModuleDeclarations.Where(p => p.DeclarationType.HasFlag(DeclarationType.Property) && p.IdentifierName == prop.IdentifierName));

        private static bool IsMoveableDeclaration(Declaration declaration)
        {
            return (declaration.DeclarationType.HasFlag(DeclarationType.Member)
                        || declaration.IsField()
                        || declaration.IsModuleConstant());
        }
    }
}
