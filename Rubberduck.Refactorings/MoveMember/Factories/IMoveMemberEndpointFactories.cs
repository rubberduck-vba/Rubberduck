using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Linq;

namespace Rubberduck.Refactorings
{
    public interface IMoveMemberEndpointFactory
    {
        IMoveMemberEndpoint CreateSourceEndpoint(Declaration target);
        IMoveMemberEndpoint CreateDestinationEndpoint(QualifiedModuleName qmn);
        IMoveMemberEndpoint CreateDestinationEndpoint(string moduleName, ComponentType componentType);
    }

    public class MoveMemberEndpointFactory : IMoveMemberEndpointFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IMoveableMemberSetsFactory _moveableMemberSetsFactory;
        public MoveMemberEndpointFactory(IDeclarationFinderProvider declarationFinderProvider, IMoveableMemberSetsFactory moveableMemberSetsFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _moveableMemberSetsFactory = moveableMemberSetsFactory;
        }

        public IMoveMemberEndpoint CreateSourceEndpoint(Declaration target)
        {
            var module = _declarationFinderProvider.DeclarationFinder.ModuleDeclaration(target.QualifiedModuleName);
            return new MoveSourceEndpoint(target, CreateMoveEndpoint(module), _moveableMemberSetsFactory);
        }

        public IMoveMemberEndpoint CreateDestinationEndpoint(QualifiedModuleName qmn)
        {
            var module = _declarationFinderProvider.DeclarationFinder.ModuleDeclaration(qmn);
            return new MoveDestinationEndpoint(CreateMoveEndpoint(module));
        }

        public IMoveMemberEndpoint CreateDestinationEndpoint(string moduleName, ComponentType componentType)
        {
            return new MoveDestinationEndpoint(CreateMoveEndpoint(moduleName, componentType));
        }

        private IMoveMemberEndpoint CreateMoveEndpoint(Declaration target)
        {
            if (target is null)
            {
                return new MoveMemberEndpoint(null, _declarationFinderProvider);
            }
            var module = _declarationFinderProvider.DeclarationFinder.ModuleDeclaration(target.QualifiedModuleName);
            return new MoveMemberEndpoint(module, _declarationFinderProvider);
        }

        private IMoveMemberEndpoint CreateMoveEndpoint(string moduleName, ComponentType componentType)
        {
            if (moduleName is null)
            {
                return new MoveMemberEndpoint(moduleName, componentType, _declarationFinderProvider);
            }

            var module = _declarationFinderProvider.DeclarationFinder.MatchName(moduleName)
                .Where(d => d.DeclarationType.HasFlag(DeclarationType.Module) && d.IsUserDefined).SingleOrDefault();

            return module != null
                ? new MoveMemberEndpoint(module, _declarationFinderProvider)
                : new MoveMemberEndpoint(moduleName, componentType, _declarationFinderProvider);
        }
    }
}
