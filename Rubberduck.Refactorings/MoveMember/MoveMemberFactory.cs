using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMoveMemberObjectsFactory
    {
        IMoveGroupsProvider CreateMoveGroupsProvider(IEnumerable<Declaration> selectedDeclarations);
        IMoveSourceModuleProxy CreateMoveSourceProxy(Declaration moveSource);
        IMoveDestinationModuleProxy CreateMoveDestinationProxy(Declaration moveDestination);
        IMoveDestinationModuleProxy CreateMoveDestination(string moduleName, ComponentType moduleComponentType = ComponentType.StandardModule);
        IMoveMemberRewriteSession CreateMoveMemberRewriteSession(IExecutableRewriteSession rewriteSession);
        IEnumerable<IMoveableMemberSet> CreateMoveables(Declaration moveTarget);
    }

    public class MoveMemberObjectsFactory : IMoveMemberObjectsFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        public MoveMemberObjectsFactory(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public IMoveGroupsProvider CreateMoveGroupsProvider(IEnumerable<Declaration> selectedDeclarations)
        {
            return new MoveGroupsProvider(selectedDeclarations, _declarationFinderProvider);
        }

        public IMoveSourceModuleProxy CreateMoveSourceProxy(Declaration target)
        {
            var sourceModule = _declarationFinderProvider.DeclarationFinder.ModuleDeclaration(target.QualifiedModuleName);
            var sourceModuleProxy = new MoveSourceModuleProxy(new MoveMemberEndpoint(sourceModule, _declarationFinderProvider));

            //sourceModuleProxy.LoadMoveableMembers(CreateMoveables(target));

            return sourceModuleProxy;
        }

        public IMoveDestinationModuleProxy CreateMoveDestinationProxy(Declaration moveDestination)
        {
            return new MoveDestinationModuleProxy(new MoveMemberEndpoint(moveDestination, _declarationFinderProvider));
        }

        public IMoveDestinationModuleProxy CreateMoveDestination(string moduleName, ComponentType moduleComponentType = ComponentType.StandardModule)
        {
            var destination = _declarationFinderProvider.DeclarationFinder.MatchName(moduleName)
                .Where(d => d.DeclarationType.HasFlag(DeclarationType.Module) && d.IsUserDefined).SingleOrDefault();

            if (destination != null)
            {
                return CreateMoveDestinationProxy(destination);
            }

            return new MoveDestinationModuleProxy(new MoveMemberEndpoint(moduleName, moduleComponentType, _declarationFinderProvider));
        }

        public IMoveMemberRewriteSession CreateMoveMemberRewriteSession(IExecutableRewriteSession rewriteSession)
        {
            return new MoveMemberRewriteSession(rewriteSession);
        }

        //TODO: Does this really go here?
        public static bool TryCreateStrategy(MoveMemberModel model, out IMoveMemberRefactoringStrategy strategy)
        {
            strategy = null;

            var strategies = new List<IMoveMemberRefactoringStrategy>();

            strategy = new MoveMemberToUndefined();
            if (strategy.IsApplicable(model))
            {
                strategies.Add(strategy);
            }

            strategy = new MoveMemberEmptySet();
            if (strategy.IsApplicable(model))
            {
                strategies.Add(strategy);
            }

            strategy = new MoveMemberToStdModule();
            if (strategy.IsApplicable(model))
            {
                strategies.Add(strategy);
            }

            if (strategies.Count() == 1)
            {
                strategy = strategies.Single();
                return true;
            }

            return false;
        }

        public IEnumerable<IMoveableMemberSet> CreateMoveables(Declaration moveTarget)
        {
            var moveableMembers = new List<IMoveableMemberSet>();
            var groupsByIdentifier = _declarationFinderProvider.DeclarationFinder.Members(moveTarget.QualifiedModuleName)
                    .Where(d => d.IsMember() || d.IsField() || d.IsModuleConstant())
                    .GroupBy(key => key.IdentifierName);

            foreach (var group in groupsByIdentifier)
            {
                var newMoveable = new MoveableMemberSet(group.ToList());
                newMoveable.IsSelected = newMoveable.IdentifierName == moveTarget.IdentifierName;
                moveableMembers.Add(newMoveable);
            }

            return moveableMembers;
        }

    }
}
