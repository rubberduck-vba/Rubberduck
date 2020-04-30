using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Collections.Generic;
using System.Linq;


namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberModel : IRefactoringModel
    {
        private readonly IMoveMemberStrategyFactory _strategyFactory;
        private readonly IMoveMemberEndpointFactory _moveEndpointFactory;

        private IMoveMemberRefactoringStrategy _strategyMoveToStandardModule;
        private IMoveMemberRefactoringStrategy _strategyMoveStdToClass;

        public MoveMemberModel(Declaration target, 
                                IDeclarationFinderProvider declarationFinderProvider, 
                                IMoveMemberStrategyFactory strategyFactory, 
                                IMoveMemberEndpointFactory moveEndpointFactory)
        {
            _moveEndpointFactory = moveEndpointFactory;

            _strategyFactory = strategyFactory;

            Source = _moveEndpointFactory.CreateSourceEndpoint(target) as IMoveSourceEndpoint;

            Destination = _moveEndpointFactory.CreateDestinationEndpoint(null, ComponentType.StandardModule) as IMoveDestinationEndpoint;

            _strategyMoveToStandardModule = _strategyFactory.Create(DetermineMoveEndpoints());
            _strategyMoveStdToClass = _strategyFactory.Create(MoveEndpoints.StdToClass);
        }

        public MoveEndpoints MoveEndpoints => DetermineMoveEndpoints();

        public IMoveSourceEndpoint Source { private set; get; }

        public IMoveDestinationEndpoint Destination { private set; get; }

        public IReadOnlyCollection<IMoveableMemberSet> MoveableMemberSets
            => Source.MoveableMembers;

        public IMoveableMemberSet MoveableMemberSetByName(string identifier)
            => Source.MoveableMemberSetByName(identifier);

        public IEnumerable<Declaration> SelectedDeclarations => MoveableMemberSets
                                            .Where(mc => mc.IsSelected)
                                            .SelectMany(selected => selected.Members);

        public void ChangeDestination(string destinationModuleName, ComponentType componentType = ComponentType.StandardModule)
        {
            Destination = _moveEndpointFactory.CreateDestinationEndpoint(destinationModuleName, componentType) as IMoveDestinationEndpoint;
        }

        public void ChangeDestination(ModuleDeclaration destinationModule)
        {
            if (destinationModule != null)
            {
                Destination = _moveEndpointFactory.CreateDestinationEndpoint(destinationModule.QualifiedModuleName) as IMoveDestinationEndpoint;
                return;
            }
            Destination = _moveEndpointFactory.CreateDestinationEndpoint(null, ComponentType.StandardModule) as IMoveDestinationEndpoint;
        }

        public bool IsExecutable
        {
            get
            {
                var result = false;
                if (TryFindApplicableStrategy(out var strategy))
                {
                    result = strategy.IsExecutableModel(this, out _);
                }
                return result;
            }
        }

        public bool TryFindApplicableStrategy(out IMoveMemberRefactoringStrategy strategy)
        {
            //The default strategy when the Destination is undefined
            if (_strategyMoveToStandardModule.IsApplicable(this))
            {
                strategy = _strategyMoveToStandardModule;
                return true;
            }
            if (_strategyMoveStdToClass.IsApplicable(this))
            {
                strategy = _strategyMoveStdToClass;
                return true;
            }
            strategy = null;
            return false;
        }

        private MoveEndpoints DetermineMoveEndpoints()
        {
            if (Source.IsStandardModule)
            {
                if (Destination.IsStandardModule)
                {
                    return MoveEndpoints.StdToStd;
                }
                return MoveEndpoints.StdToClass;
            }
            if (Source.IsClassModule)
            {
                if (Destination.IsStandardModule)
                {
                    return MoveEndpoints.ClassToStd;
                }
                return MoveEndpoints.ClassToClass;
            }
            if (Source.IsUserFormModule)
            {
                if (Destination.IsStandardModule)
                {
                    return MoveEndpoints.FormToStd;
                }
                return MoveEndpoints.FormToClass;
            }
            return MoveEndpoints.Undefined;
        }

    }
}
