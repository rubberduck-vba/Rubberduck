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
    public enum MoveGroups
    {
        AllParticipants,
        Selected, 
        Support,
        NonParticipants,
        Support_Public
    }

    public interface IMoveGroupsProvider
    {
        /// <summary>
        /// Returns read-only declarations collection for the moveGroup
        /// </summary>
        IReadOnlyCollection<Declaration> Declarations(MoveGroups moveGroup);

        /// <summary>
        /// Returns read-only dependency declarations collection for the moveGroup
        /// </summary>
        IReadOnlyCollection<Declaration> Dependencies(MoveGroups moveGroup);

        /// <summary>
        /// Returns IMoveableMemberSet for all moveable Source Module declarations
        /// </summary>
        IReadOnlyCollection<IMoveableMemberSet> MoveableMembers { get; }

        /// <summary>
        /// Returns IMoveableMemberSet for the identifier
        /// </summary>
        IMoveableMemberSet MoveableMember(string identifier);

        /// <summary>
        /// Returns IMoveableMemberSet for the specified MoveGroup
        /// </summary>
        IReadOnlyCollection<IMoveableMemberSet> this[MoveGroups moveGroup] { get; }

        /// <summary>
        /// Returns all declarations involved in the MoveMember refactoring request 
        /// </summary>
        IReadOnlyCollection<Declaration> AllParticipants { get; }

        /// <summary>
        /// Returns the explicitly selected Declarations that define the move 
        /// </summary>
        IReadOnlyCollection<Declaration> Selected { get; }

        /// <summary>
        /// Returns declaration that are referenced directly or indirectly by 
        /// the MoveMember CallTree declarations
        /// </summary>
        IReadOnlyCollection<Declaration> SupportMembers { get; }

        /// <summary>
        /// Returns the supporting declarations referenced exclusively by the MoveMember calltrees 
        /// </summary>
        IReadOnlyCollection<Declaration> ExclusiveSupportDeclarations { get; }
        
        /// <summary>
        /// Returns MoveMember support declarations that are referenced by
        /// Source module members not involved in the move 
        /// </summary>
        IReadOnlyCollection<Declaration> NonExclusiveSupportDeclarations { get; }

        IReadOnlyCollection<Declaration> FlattenedDependencyGraph(string identifier);
    }

    public class MoveGroupsProvider : IMoveGroupsProvider
    {
        private readonly IDeclarationFinderProvider _declarationProvider;
        private List<IMoveableMemberSet> _allMoveableMemberSets;

        public MoveGroupsProvider(IEnumerable<IMoveableMemberSet> moveableMemberSets, IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationProvider = declarationFinderProvider;
            _allMoveableMemberSets = moveableMemberSets.ToList();
            foreach (var moveableMemberSet in moveableMemberSets)
            {
                moveableMemberSet.IsSupport = false;
                moveableMemberSet.IsExclusive = false;
            }

            var selectedDeclarations = moveableMemberSets.Where(mm => mm.IsSelected).SelectMany(mm => mm.Members);

            if (!selectedDeclarations.Any()) { return; }

            var allParticipants = CallTreeDeclarations(selectedDeclarations, declarationFinderProvider);

            foreach (var supportParticipant in allParticipants.Except(selectedDeclarations))
            {
                var moveMemberSet = moveableMemberSets.Where(mm => mm.Contains(supportParticipant)).SingleOrDefault();
                if (moveMemberSet != null)
                {
                    moveMemberSet.IsSupport = true;
                    moveMemberSet.IsExclusive = allParticipants.ContainsParentScopesForAllReferences(moveMemberSet.NonMemberBodyReferences);
                }
            }

            foreach (var moveable in _allMoveableMemberSets)
            {
                var fdg = BuildFlattenedDependencyGraph(moveable.IdentifierName);
                moveable.FlattenedDependencyGraph = fdg;
            }
        }

        public IMoveableMemberSet this[string identifier]
        {
            get
            {
                return _allMoveableMemberSets.Where(mm => mm.IdentifierName.IsEquivalentVBAIdentifierTo(identifier))
                    .SingleOrDefault();
            }
        }

        public IReadOnlyCollection<IMoveableMemberSet> this[MoveGroups moveGroup]
        {
            get
            {
                switch (moveGroup)
                {
                    case MoveGroups.Selected:
                        return _allMoveableMemberSets.Where(mm => mm.IsSelected).ToList();
                    case MoveGroups.AllParticipants:
                        return _allMoveableMemberSets.Where(mm => mm.IsSelected || mm.IsSupport).ToList();
                    case MoveGroups.Support:
                        return _allMoveableMemberSets.Where(mm => !mm.IsSelected && mm.IsSupport).ToList();
                    case MoveGroups.Support_Public:
                        return _allMoveableMemberSets.Where(mm => !mm.HasPrivateAccessibility && (!mm.IsSelected && mm.IsSupport)).ToList();
                    case MoveGroups.NonParticipants:
                        return _allMoveableMemberSets.Where(mm => !(mm.IsSelected || mm.IsSupport)).ToList();
                }
                return new List<IMoveableMemberSet>();
            }
        }

        public IReadOnlyCollection<Declaration> Declarations(MoveGroups moveGroup)
        {
            switch (moveGroup)
            {
                case MoveGroups.Selected:
                    return Selected;
                case MoveGroups.AllParticipants:
                    return AllParticipants;
                case MoveGroups.Support:
                    return Support;
                case MoveGroups.Support_Public:
                    return Support.Where(d => !d.HasPrivateAccessibility()).ToList();
                case MoveGroups.NonParticipants:
                    return _allMoveableMemberSets.Where(mm => !(mm.IsSelected || mm.IsSupport))
                                                    .SelectMany(mm => mm.Members)
                                                    .ToList();
            }
            return new List<Declaration>();
        }

        public IReadOnlyCollection<Declaration> Dependencies(MoveGroups moveGroup)
        {
            switch (moveGroup)
            {
                case MoveGroups.Selected:
                    return AggregateDependencies(this[MoveGroups.Selected]);
                case MoveGroups.AllParticipants:
                    return AggregateDependencies(this[MoveGroups.AllParticipants]);
                case MoveGroups.Support:
                    return AggregateDependencies(this[MoveGroups.Support]);
                case MoveGroups.NonParticipants:
                    return AggregateDependencies(this[MoveGroups.NonParticipants]);
                case MoveGroups.Support_Public:
                    return AggregateDependencies(this[MoveGroups.Support].Where(mm => !mm.HasPrivateAccessibility));
            }
            return new List<Declaration>();
        }

        private IReadOnlyCollection<Declaration> AggregateDependencies(IEnumerable<IMoveableMemberSet> moveMemberSets )
        {
            var aggregated = new List<Declaration>();
            foreach (var moveMemberSet in moveMemberSets)
            {
                var dependencies = FlattenedDependencyGraph(moveMemberSet.IdentifierName).ToList();
                aggregated.AddRange(dependencies);
            }
            return aggregated;
        }

        public IReadOnlyCollection<Declaration> FlattenedDependencyGraph(string identifier)
        {
            return this[identifier].FlattenedDependencyGraph;
        }

        private List<Declaration> BuildFlattenedDependencyGraph(string identifier)
        {
            var names = new List<string>();

            var moveable = this[identifier];
            var dependentNames = moveable.DirectDependencies.Select(d => d.IdentifierName).ToList();

            while (TryAddDependencyNames(names, dependentNames, out var newNames))
            {
                dependentNames = newNames;
            }

            var flattened = new List<Declaration>();
            foreach (var id in names)
            {
                flattened.AddRange(this[id].Members);
            }
            return flattened;
        }

        private bool TryAddDependencyNames(List<string> names, List<string> dependencyNames, out List<string> newNames)
        {
            var newDependencyNames = new List<string>();
            if (dependencyNames.Any())
            {
                foreach (var name in dependencyNames)
                {
                    var moveable = this[name];
                    if (moveable != null) //UDT members return null
                    {
                        names.Add(name);
                        newDependencyNames.AddRange(moveable.DirectDependencies.Select(d => d.IdentifierName));
                    }
                }
            }
            newNames = newDependencyNames;
            return dependencyNames.Any();
        }

        public IMoveableMemberSet MoveableMember(string identifier)
        {
            return this[identifier];
        }

        public IReadOnlyCollection<IMoveableMemberSet> MoveableMembers => _allMoveableMemberSets;

        public IReadOnlyCollection<Declaration> Selected 
            => _allMoveableMemberSets.Where(mm => mm.IsSelected)
                                    .SelectMany(mm => mm.Members)
                                    .ToList();

        public IReadOnlyCollection<Declaration> AllParticipants 
            => _allMoveableMemberSets.Where(mm => mm.IsSelected || mm.IsSupport)
                                    .SelectMany(mm => mm.Members)
                                    .ToList();

        public IReadOnlyCollection<Declaration> ExclusiveSupportDeclarations 
            => _allMoveableMemberSets.Where(mm => mm.IsSupport && mm.IsExclusive)
                                    .SelectMany(mm => mm.Members)
                                    .ToList();

        public IReadOnlyCollection<Declaration> NonExclusiveSupportDeclarations 
            => _allMoveableMemberSets.Where(mm => mm.IsSupport && !mm.IsExclusive)
                                    .SelectMany(mm => mm.Members)
                                    .ToList();

        public IReadOnlyCollection<Declaration> Support
            => _allMoveableMemberSets.Where(mm => mm.IsSupport)
                                    .SelectMany(mm => mm.Members)
                                    .ToList();

        public IReadOnlyCollection<Declaration> SupportMembers
            => _allMoveableMemberSets.Where(mm => mm.IsSupport && mm.Member.IsMember())
                                    .SelectMany(mm => mm.Members)
                                    .ToList();

        private static IReadOnlyCollection<Declaration> CallTreeDeclarations(IEnumerable<Declaration> definingDeclarations, IDeclarationFinderProvider declarationFinderProvider)
        {
            if (!definingDeclarations.Any()) { return new HashSet<Declaration>(); }

            var allElements = declarationFinderProvider.DeclarationFinder.Members(definingDeclarations.First().QualifiedModuleName)
                .Where(d => !d.DeclarationType.HasFlag(DeclarationType.Module));

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

            return participatingDeclarations.ToList();
        }

        private static bool IsMoveableDeclaration(Declaration declaration)
        {
            return (declaration.DeclarationType.HasFlag(DeclarationType.Member)
                        || declaration.IsField()
                        || declaration.IsModuleConstant());
        }
    }
}
