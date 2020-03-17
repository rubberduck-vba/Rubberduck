using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember
{
    public enum MoveGroup
    {
        AllParticipants,
        Selected, 
        Support,
        NonParticipants,
        Support_Public,
        Support_Private,
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
        /// Returns read-only IMoveableMemberSet collection for the specified MoveGroup
        /// </summary>
        IReadOnlyCollection<IMoveableMemberSet> MoveableMemberSets(MoveGroup moveGroup);

        /// <summary>
        /// Returns read-only flattened dependency graph for a group of IMoveMemberSets
        /// </summary>
        IReadOnlyCollection<Declaration> AggregateDependencies(IEnumerable<IMoveableMemberSet> moveMemberSets);

        /// <summary>
        /// Returns a collection of MoveMemberSets for a group of declarations
        /// </summary>
        IReadOnlyCollection<IMoveableMemberSet> ToMoveableMemberSets(IEnumerable<Declaration> declarations);
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
            _allParticipants = new List<Declaration>();

            CreateFlattenedDependencies(_allMoveableMemberSets);

            _allParticipants =
                selectedDeclarations.Concat(AggregateDependencies(selectedMoveMemberSets)).ToList();

            foreach (var moveableMemberSet in _allMoveableMemberSets)
            {
                moveableMemberSet.IsSupport = !moveableMemberSet.IsSelected 
                                                    && (_allParticipants.Contains(moveableMemberSet.Member) 
                                                        || _allParticipants.Any(p => (p.AsTypeDeclaration?.IdentifierName.Equals(moveableMemberSet.Member.IdentifierName) ?? false)));

                if (moveableMemberSet.IsSupport)
                {
                    SetIsExclusiveFlag(moveableMemberSet);
                }
            }

            _moveMemberSetsByMoveGroup = new Dictionary<MoveGroup, List<IMoveableMemberSet>>()
            {
                [MoveGroup.Selected] = _allMoveableMemberSets.Where(mm => mm.IsSelected).ToList(),
                [MoveGroup.AllParticipants] = _allMoveableMemberSets.Where(mm => mm.IsSelected || mm.IsSupport).ToList(),
                [MoveGroup.NonParticipants] = _allMoveableMemberSets.Where(mm => !(mm.IsSelected || mm.IsSupport)).ToList(),
                [MoveGroup.Support] = _allMoveableMemberSets.Where(mm => !mm.IsSelected && mm.IsSupport).ToList(),
                [MoveGroup.Support_Public] = _allMoveableMemberSets.Where(mm => !mm.HasPrivateAccessibility && (!mm.IsSelected && mm.IsSupport)).ToList(),
                [MoveGroup.Support_Private] = _allMoveableMemberSets.Where(mm => mm.HasPrivateAccessibility && (!mm.IsSelected && mm.IsSupport)).ToList(),
                [MoveGroup.Support_Exclusive] = _allMoveableMemberSets.Where(mm => !mm.IsSelected && mm.IsSupport && mm.IsExclusive).ToList(),
                [MoveGroup.Support_NonExclusive] = _allMoveableMemberSets.Where(mm => !mm.IsSelected && mm.IsSupport && !mm.IsExclusive).ToList(),
            };

            _declarationsByMoveGroup = new Dictionary<MoveGroup, List<Declaration>>()
            {
                [MoveGroup.AllParticipants] = _allParticipants.ToList(),
                [MoveGroup.NonParticipants] = _allMoveableMemberSets
                                                    .Where(mm => !(mm.IsSelected || mm.IsSupport))
                                                    .SelectMany(mm => mm.Members)
                                                    .ToList(),
                [MoveGroup.Selected] = _moveMemberSetsByMoveGroup[MoveGroup.Selected]
                                                    .SelectMany(mm => mm.Members)
                                                    .ToList(),
                [MoveGroup.Support] = _moveMemberSetsByMoveGroup[MoveGroup.Support]
                                                    .SelectMany(mm => mm.Members)
                                                    .ToList(),
                [MoveGroup.Support_Public] = _moveMemberSetsByMoveGroup[MoveGroup.Support]
                                                    .Where(mm => !mm.HasPrivateAccessibility)
                                                    .SelectMany(mm => mm.Members)
                                                    .ToList(),
                [MoveGroup.Support_Private] = _moveMemberSetsByMoveGroup[MoveGroup.Support]
                                                    .Where(mm => mm.HasPrivateAccessibility)
                                                    .SelectMany(mm => mm.Members)
                                                    .ToList(),
                [MoveGroup.Support_Exclusive] = _moveMemberSetsByMoveGroup[MoveGroup.Support_Exclusive]
                                                    .SelectMany(mm => mm.Members)
                                                    .ToList(),
                [MoveGroup.Support_NonExclusive] = _moveMemberSetsByMoveGroup[MoveGroup.Support_NonExclusive]
                                                    .SelectMany(mm => mm.Members)
                                                    .ToList()
            };

            if (_declarationsByMoveGroup[MoveGroup.AllParticipants].Intersect(_declarationsByMoveGroup[MoveGroup.NonParticipants]).Any())
            {
                throw new MoveMemberUnsupportedMoveException(moveableMemberSets.Select(mm => mm.Member).FirstOrDefault());
            }

            _dependenciesByMoveGroup = new Dictionary<MoveGroup, List<Declaration>>();
        }

        public IReadOnlyCollection<IMoveableMemberSet> MoveableMemberSets(MoveGroup moveGroup) 
            => _moveMemberSetsByMoveGroup is null 
                ? new List<IMoveableMemberSet>() 
                : _moveMemberSetsByMoveGroup[moveGroup];

        public IReadOnlyCollection<Declaration> Declarations(MoveGroup moveGroup) 
            => _declarationsByMoveGroup is null 
                    ? new List<Declaration>() 
                    : _declarationsByMoveGroup[moveGroup];

        public IReadOnlyCollection<Declaration> Dependencies(MoveGroup moveGroup)
        {
            if (_dependenciesByMoveGroup is null)
            {
                return new List<Declaration>();
            }

            if (!_dependenciesByMoveGroup.TryGetValue(moveGroup, out var dependencies))
            {
                dependencies = AggregateDependencies(MoveableMemberSets(moveGroup)).ToList();
                _dependenciesByMoveGroup.Add(moveGroup, dependencies);
            }
            return dependencies;
        }

        public IReadOnlyCollection<IMoveableMemberSet> ToMoveableMemberSets(IEnumerable<Declaration> declarations)
        {
            var uniqueIdentifiers = declarations.Select(d => d.IdentifierName).Distinct();
            var moveables = new List<IMoveableMemberSet>();
            foreach (var identifier in uniqueIdentifiers)
            {
                moveables.AddRange(_moveMemberSetsByMoveGroup[MoveGroup.AllParticipants].Where(mm => mm.IdentifierName.IsEquivalentVBAIdentifierTo(identifier)));
            }
            return moveables;
        }

        private IMoveableMemberSet MoveableMemberSet(Declaration declaration)
        {
            if (_moveableMembersByName.TryGetValue(declaration.IdentifierName, out var moveable)
                && moveable.Members.Any(mm => mm.DeclarationType.Equals(declaration.DeclarationType)))
            {
                return moveable;
            }
            return null;
        }

        public IReadOnlyCollection<Declaration> AggregateDependencies(IEnumerable<IMoveableMemberSet> moveMemberSets )
        {
            var aggregated = new List<Declaration>();
            foreach (var moveMemberSet in moveMemberSets)
            {
                var dependencies = MoveableMemberSet(moveMemberSet.Member).FlattenedDependencies.ToList();
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
                foreach (var declaration in dependencyDeclarations)
                {
                    flattened.AddRange(MoveableMemberSet(declaration).Members);
                }
                moveable.FlattenedDependencies = flattened;
            }
        }

        private List<Declaration> AddDependencies(List<Declaration> allDependencies, List<Declaration> dependenciesToAdd, out List<Declaration> downstreamDependencies)
        {
            downstreamDependencies = new List<Declaration>();
            foreach (var dependency in dependenciesToAdd)
            {
                var moveable = MoveableMemberSet(dependency);
                if (moveable is null) //e.g., UserDefinedTypeMember results in null
                {
                    continue;
                }
                allDependencies.Add(dependency);
                downstreamDependencies.AddRange(moveable.DirectDependencies);
            }
            return allDependencies;
        }

        private void SetIsExclusiveFlag(IMoveableMemberSet moveableMemberSet)
        {
            var referencesExternalToMember = moveableMemberSet.Members.AllReferences().Where(rf => !moveableMemberSet.Members.Contains(rf.ParentScoping));

            moveableMemberSet.IsExclusive = _allParticipants.ContainsParentScopesForAll(referencesExternalToMember);

            if (!moveableMemberSet.IsExclusive)
            {
                var qmnSource = moveableMemberSet.Member.QualifiedModuleName;
                var participatingTypeFields = _allParticipants.Where(p => p.IsField() && p.AsTypeName.Equals(moveableMemberSet.Member.IdentifierName));
                if (moveableMemberSet.IsUserDefinedType)
                {
                    var allUDTFields = _declarationProvider.DeclarationFinder.Members(qmnSource).Where(m => m.IsField() && m.AsTypeDeclaration.DeclarationType.Equals(DeclarationType.UserDefinedType));
                    moveableMemberSet.IsExclusive = !(allUDTFields.Except(participatingTypeFields)).Any();
                }
                if (moveableMemberSet.IsEnumeration)
                {
                    var allEnumFields = _declarationProvider.DeclarationFinder.Members(qmnSource).Where(m => m.IsField() && m.AsTypeDeclaration.DeclarationType.Equals(DeclarationType.Enumeration));
                    moveableMemberSet.IsExclusive = !(allEnumFields.Except(participatingTypeFields)).Any();
                }
            }
        }
    }
}
