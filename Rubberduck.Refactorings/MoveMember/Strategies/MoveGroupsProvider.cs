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
    public enum MoveGroup
    {
        AllParticipants,
        Selected, 
        Support,
        NonParticipants,
        Support_Public,
        Support_Exclusive,
        Support_NonExclusive
    }

    public interface IMoveGroupsProvider
    {
        /// <summary>
        /// Returns read-only declarations collection for the moveGroup
        /// </summary>
        IReadOnlyCollection<Declaration> Declarations(MoveGroup moveGroup);

        /// <summary>
        /// Returns read-only dependency declarations collection for the moveGroup
        /// </summary>
        IReadOnlyCollection<Declaration> Dependencies(MoveGroup moveGroup);

        /// <summary>
        /// Returns IMoveableMemberSet for the specified MoveGroup
        /// </summary>
        IReadOnlyCollection<IMoveableMemberSet> MoveableMemberSets(MoveGroup moveGroup);

    }

    public class MoveGroupsProvider : IMoveGroupsProvider
    {
        private readonly IDeclarationFinderProvider _declarationProvider;
        private List<IMoveableMemberSet> _allMoveableMemberSets;

        private List<Declaration> _allParticipants;
        private Dictionary<MoveGroup, List<Declaration>> _declarationsByMoveGroup;
        private Dictionary<MoveGroup, List<IMoveableMemberSet>> _moveMemberSetsByMoveGroup;
        private Dictionary<MoveGroup, List<Declaration>> _dependenciesByMoveGroup;
        private Dictionary<string, IMoveableMemberSet> _moveableMembersByName;

        public MoveGroupsProvider(IEnumerable<IMoveableMemberSet> moveableMemberSets, IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationProvider = declarationFinderProvider;
            _allMoveableMemberSets = moveableMemberSets.ToList();

            var selectedMoveMemberSets = _allMoveableMemberSets.Where(mm => mm.IsSelected);
            var selectedDeclarations = selectedMoveMemberSets.SelectMany(mm => mm.Members);

            if (!selectedDeclarations.Any()) { return; }

            _moveableMembersByName = _allMoveableMemberSets.ToDictionary(key => key.IdentifierName);

            CreateFlattenedDependencies(_allMoveableMemberSets);

            _allParticipants =
                selectedDeclarations.Concat(AggregateDependencies(selectedMoveMemberSets)).ToList();

            foreach (var moveableMemberSet in _allMoveableMemberSets)
            {
                moveableMemberSet.IsSupport = !moveableMemberSet.IsSelected && _allParticipants.Contains(moveableMemberSet.Member);

                var referencesExternalToMember = moveableMemberSet.Members.AllReferences().Where(rf => !moveableMemberSet.Members.Contains(rf.ParentScoping));
                moveableMemberSet.IsExclusive = _allParticipants.ContainsParentScopesForAll(referencesExternalToMember);
            }

            _declarationsByMoveGroup = new Dictionary<MoveGroup, List<Declaration>>()
            {
                [MoveGroup.AllParticipants] = _allParticipants.ToList(),
                [MoveGroup.NonParticipants] = _allMoveableMemberSets
                                                    .Where(mm => !(mm.IsSelected || mm.IsSupport))
                                                    .SelectMany(mm => mm.Members)
                                                    .ToList(),
                [MoveGroup.Selected] = _allMoveableMemberSets
                                                    .Where(mm => mm.IsSelected)
                                                    .SelectMany(mm => mm.Members)
                                                    .ToList(),
                [MoveGroup.Support] = _allMoveableMemberSets
                                                    .Where(mm => mm.IsSupport && !mm.IsSelected)
                                                    .SelectMany(mm => mm.Members)
                                                    .ToList(),
                [MoveGroup.Support_Public] = _allMoveableMemberSets
                                                    .Where(mm => mm.IsSupport && !mm.IsSelected && !mm.HasPrivateAccessibility)
                                                    .SelectMany(mm => mm.Members)
                                                    .ToList(),
                [MoveGroup.Support_Exclusive] = _allMoveableMemberSets
                                                    .Where(mm => mm.IsSupport && !mm.IsSelected && mm.IsExclusive)
                                                    .SelectMany(mm => mm.Members)
                                                    .ToList(),
                [MoveGroup.Support_NonExclusive] = _allMoveableMemberSets
                                                    .Where(mm => mm.IsSupport && !mm.IsSelected && !mm.IsExclusive)
                                                    .SelectMany(mm => mm.Members)
                                                    .ToList()
            };

            _moveMemberSetsByMoveGroup = new Dictionary<MoveGroup, List<IMoveableMemberSet>>()
            {
                [MoveGroup.Selected] = _allMoveableMemberSets.Where(mm => mm.IsSelected).ToList(),
                [MoveGroup.AllParticipants] = _allMoveableMemberSets.Where(mm => mm.IsSelected || mm.IsSupport).ToList(),
                [MoveGroup.NonParticipants] = _allMoveableMemberSets.Where(mm => !(mm.IsSelected || mm.IsSupport)).ToList(),
                [MoveGroup.Support] = _allMoveableMemberSets.Where(mm => !mm.IsSelected && mm.IsSupport).ToList(),
                [MoveGroup.Support_Public] = _allMoveableMemberSets.Where(mm => !mm.HasPrivateAccessibility && (!mm.IsSelected && mm.IsSupport)).ToList(),
                [MoveGroup.Support_Exclusive] = _allMoveableMemberSets.Where(mm => !mm.IsSelected && mm.IsSupport && mm.IsExclusive).ToList(),
                [MoveGroup.Support_NonExclusive] = _allMoveableMemberSets.Where(mm => !mm.IsSelected && mm.IsSupport && !mm.IsExclusive).ToList(),
            };

            _dependenciesByMoveGroup = new Dictionary<MoveGroup, List<Declaration>>();
        }

        public IReadOnlyCollection<IMoveableMemberSet> MoveableMemberSets(MoveGroup moveGroup) 
            => _moveMemberSetsByMoveGroup[moveGroup];

        public IReadOnlyCollection<Declaration> Declarations(MoveGroup moveGroup) 
            => _declarationsByMoveGroup[moveGroup];

        public IReadOnlyCollection<Declaration> Dependencies(MoveGroup moveGroup)
        {
            if (!_dependenciesByMoveGroup.TryGetValue(moveGroup, out var dependencies))
            {
                dependencies = AggregateDependencies(MoveableMemberSets(moveGroup)).ToList();
                _dependenciesByMoveGroup.Add(moveGroup, dependencies);
            }
            return dependencies;
        }

        private IMoveableMemberSet MoveableMemberSet(string identifier)
            =>  _moveableMembersByName.TryGetValue(identifier, out var moveable)
                    ? moveable
                    : null;

        private IReadOnlyCollection<Declaration> AggregateDependencies(IEnumerable<IMoveableMemberSet> moveMemberSets )
        {
            var aggregated = new List<Declaration>();
            foreach (var moveMemberSet in moveMemberSets)
            {
                var dependencies = MoveableMemberSet(moveMemberSet.IdentifierName).FlattenedDependencies.ToList();
                aggregated.AddRange(dependencies);
            }
            return aggregated;
        }

        private void CreateFlattenedDependencies(IEnumerable<IMoveableMemberSet> moveableMembers)
        {
            foreach (var moveable in moveableMembers)
            {
                var dependencyDeclarations = new List<Declaration>();

                var dependencies = moveable.DirectDependencies.ToList();

                while(dependencies.Any())
                {
                    dependencyDeclarations = AddDependencies(dependencyDeclarations, dependencies, out var additionalDependencies);
                    dependencies = additionalDependencies;
                }

                //Once all the dependencies by declaration are found,
                //we want a flattened list of dependencies that includes 
                //all the declarations by MoveableMemberSet.  
                //This attaches the related Property members as dependencies (for the purposes of moving)
                // even if only one of them participates in a given dependency graph
                var flattened = new List<Declaration>();
                foreach (var id in dependencyDeclarations)
                {
                    flattened.AddRange(MoveableMemberSet(id.IdentifierName).Members);
                }
                moveable.FlattenedDependencies = flattened;
            }
        }

        private List<Declaration> AddDependencies(List<Declaration> allDependencies, List<Declaration> dependenciesToAdd, out List<Declaration> downstreamDependencies)
        {
            downstreamDependencies = new List<Declaration>();
            foreach (var dependency in dependenciesToAdd)
            {
                var moveable = MoveableMemberSet(dependency.IdentifierName);
                if (moveable != null) //e.g., UDT members return null
                {
                    allDependencies.Add(dependency);
                    downstreamDependencies.AddRange(moveable.DirectDependencies);
                }
            }
            return allDependencies;
        }
    }
}
