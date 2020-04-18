using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings
{
    public interface IMoveMemberModelFactory
    {
        MoveMemberModel Create(Declaration target);
        MoveMemberModel Create(Declaration target, ModuleDeclaration destination);
        MoveMemberModel Create(Declaration target, string destinationModuleName, DeclarationType declarationType = DeclarationType.ProceduralModule);
        MoveMemberModel Create(IEnumerable<Declaration> targets);
        MoveMemberModel Create(IEnumerable<Declaration> targets, ModuleDeclaration destination);
        MoveMemberModel Create(IEnumerable<Declaration> targets, string destinationModuleName, DeclarationType declarationType = DeclarationType.ProceduralModule);
    }

    public class MoveMemberModelFactory : IMoveMemberModelFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IMoveMemberStrategyFactory _strategyFactory;
        private readonly IMoveMemberEndpointFactory _endpointFactory;

        public MoveMemberModelFactory(IDeclarationFinderProvider declarationFinderProvider,
                                        IMoveMemberStrategyFactory strategyFactory,
                                        IMoveMemberEndpointFactory endpointFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _strategyFactory = strategyFactory;
            _endpointFactory = endpointFactory;
        }

        public MoveMemberModel Create(Declaration target)
        {
            if (target == null) { throw new TargetDeclarationIsNullException(); }

            ThrowIfInvalidTargetSet(target);
            return new MoveMemberModel(target, _declarationFinderProvider, _strategyFactory, _endpointFactory);
        }

        public MoveMemberModel Create(Declaration target, ModuleDeclaration destination)
        {
            if (destination == null) { throw new TargetDeclarationIsNullException(); }

            ThrowIfInvalidTargetSet(target);
            var model = new MoveMemberModel(target, _declarationFinderProvider, _strategyFactory, _endpointFactory);
            model.ChangeDestination(destination);
            return model;
        }

        public MoveMemberModel Create(Declaration target, string destinationModuleName, DeclarationType declarationType = DeclarationType.ProceduralModule)
        {
            var model = new MoveMemberModel(target, _declarationFinderProvider, _strategyFactory, _endpointFactory);
            var componenType = DeclarationTypeToComponentType(declarationType);

            model.ChangeDestination(destinationModuleName, componenType);
            return model;
        }

        public MoveMemberModel Create(IEnumerable<Declaration> targets)
        {
            ThrowIfInvalidTargetSet(targets.ToArray());

            var model = Create(targets.First());
            return SelectAllTargets(model, targets);
        }

        public MoveMemberModel Create(IEnumerable<Declaration> targets, ModuleDeclaration destination)
        {
            ThrowIfInvalidTargetSet(targets.ToArray());

            var model = Create(targets.First(), destination);
            return SelectAllTargets(model, targets);
        }

        public MoveMemberModel Create(IEnumerable<Declaration> targets, string destinationModuleName, DeclarationType declarationType = DeclarationType.ProceduralModule)
        {
            ThrowIfInvalidTargetSet(targets.ToArray());

            var model = Create(targets.First(), destinationModuleName, declarationType);
            return SelectAllTargets(model, targets);
        }

        private MoveMemberModel SelectAllTargets(MoveMemberModel model, IEnumerable<Declaration> targets)
        {
            foreach (var target in targets)
            {
                model.MoveableMemberSetByName(target.IdentifierName).IsSelected = true;
            }
            return model;
        }

        private void ThrowIfInvalidTargetSet(params Declaration[] targets)
        {
            if (!targets.All(t => t.IsMember()
                                    || t.IsModuleConstant()
                                    || t.IsMemberVariable()))
            {
                throw new MoveMemberUnsupportedMoveException();
            }
        }

        private ComponentType DeclarationTypeToComponentType(DeclarationType declarationType)
        {
            return declarationType.Equals(DeclarationType.ProceduralModule)
                ? ComponentType.StandardModule
                : ComponentType.ClassModule;
        }
    }
}
