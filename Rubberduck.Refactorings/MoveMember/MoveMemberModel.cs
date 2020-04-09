using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Linq;


namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberModel : IRefactoringModel
    {
        private readonly IMoveMemberStrategyFactory _strategyFactory;
        private readonly IMoveMemberEndpointFactory _moveEndpointFactory;

        private IMoveMemberRefactoringStrategy _strategyMoveToStandardModule;

        public IDeclarationFinderProvider DeclarationFinderProvider { get; }

        public MoveMemberModel(Declaration target, 
                                IDeclarationFinderProvider declarationFinderProvider, 
                                IMoveMemberStrategyFactory strategyFactory, 
                                IMoveMemberEndpointFactory moveEndpointFactory)
        {
            DeclarationFinderProvider = declarationFinderProvider;

            _moveEndpointFactory = moveEndpointFactory;

            _strategyFactory = strategyFactory;

            Source = _moveEndpointFactory.CreateSourceEndpoint(target) as IMoveSourceEndpoint;

            var destinationModuleName = DetermineInitialDestinationModuleName(declarationFinderProvider, Source.ModuleName);
            Destination = _moveEndpointFactory.CreateDestinationEndpoint(destinationModuleName, ComponentType.StandardModule) as IMoveDestinationEndpoint;

            _strategyMoveToStandardModule = _strategyFactory.Create(MoveMemberStrategy.MoveToStandardModule);
        }

        public IMoveSourceEndpoint Source { private set; get; }

        public IMoveDestinationEndpoint Destination { private set; get; }

        public IReadOnlyCollection<IMoveableMemberSet> MoveableMembers
            => Source.MoveableMembers;

        public IMoveableMemberSet MoveableMemberSetByName(string identifier)
            => Source.MoveableMemberSetByName(identifier);

        public IEnumerable<Declaration> SelectedDeclarations => MoveableMembers
                                            .Where(mc => mc.IsSelected)
                                            .SelectMany(selected => selected.Members);

        public void ChangeDestination(string destinationModuleName, ComponentType componentType = ComponentType.StandardModule)
        {
            Destination = _moveEndpointFactory.CreateDestinationEndpoint(destinationModuleName, componentType) as IMoveDestinationEndpoint;
        }

        public void ChangeDestination(Declaration destinationModule)
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
            strategy = null;
            return false;
        }

        public IMoveMemberRefactoringPreviewerFactory PreviewerFactory { set; get; }

        public bool TryGetPreview(IMoveMemberEndpoint endpoint, out string preview)
        {
            var previewer = PreviewerFactory?.Create(endpoint);
            preview = previewer?.PreviewMove(this) ?? string.Empty;
            return previewer != null;
        }

        //public bool TryGetPreviewerFactory(out IMoveMemberRefactoringPreviewerFactory factory)
        //{
        //    factory = PreviewerFactory;
        //    return factory != null;
        //}

        private static string DetermineInitialDestinationModuleName(IDeclarationFinderProvider declarationFinderProvider, string sourceModuleName)
        {
            var allModuleIdentifiers = declarationFinderProvider.DeclarationFinder.AllModules.Select(m => m.ComponentName);
            var destinationModuleName = sourceModuleName;
            var hasNameConflict = true;
            for (var idx = 0; hasNameConflict && idx < 100; idx++)
            {
                destinationModuleName = destinationModuleName.IncrementIdentifier();
                if (allModuleIdentifiers.All(name => !destinationModuleName.IsEquivalentVBAIdentifierTo(name)))
                {
                    hasNameConflict = false;
                }
            }
            return destinationModuleName;
        }
    }
}
