using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMoveMemberObjectsFactory
    {
        MoveGroupsProvider CreateMoveGroupsProvider(IEnumerable<IMoveableMemberSet> selectedDeclarations);
        IMoveSourceModuleProxy CreateMoveSourceProxy(Declaration moveSource);
        IMoveDestinationModuleProxy CreateMoveDestinationProxy(Declaration moveDestination);
        IMoveDestinationModuleProxy CreateMoveDestination(string moduleName, ComponentType moduleComponentType = ComponentType.StandardModule);
        IEnumerable<IMoveableMemberSet> CreateMoveables(Declaration moveTarget);
        IMovedContentProvider CreateMovedContentProvider();
    }

    public class MoveMemberObjectsFactory : IMoveMemberObjectsFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        public MoveMemberObjectsFactory(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public MoveGroupsProvider CreateMoveGroupsProvider(IEnumerable<IMoveableMemberSet> moveableMemberSets)
        {
            return new MoveGroupsProvider(moveableMemberSets, _declarationFinderProvider);
        }

        public IMoveSourceModuleProxy CreateMoveSourceProxy(Declaration target)
        {
            var sourceModule = _declarationFinderProvider.DeclarationFinder.ModuleDeclaration(target.QualifiedModuleName);
            return new MoveSourceModuleProxy(new MoveMemberEndpoint(sourceModule, _declarationFinderProvider));
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

        public static bool TryCreateStrategy(MoveMemberModel model, out IMoveMemberRefactoringStrategy strategy)
        {
            strategy = null;

            var strategies = new List<IMoveMemberRefactoringStrategy>();

            strategy = new MoveMemberEmptySet();
            if (strategy.IsApplicable(model))
            {
                strategies.Add(strategy);
            }

            //The default strategy when the Destination is undefined
            strategy = new MoveMemberToStdModule();
            if (strategy.IsApplicable(model))
            {
                strategies.Add(strategy);
            }

            //Unless a single applicable strategies is found,
            //the correct strategy is indeterminant.
            if (strategies.Count() == 1)
            {
                strategy = strategies.Single();
                return true;
            }

            return false;
        }

        public IEnumerable<IMoveableMemberSet> CreateMoveables(Declaration moveTarget)
        {
            var groupsByIdentifier = _declarationFinderProvider.DeclarationFinder.Members(moveTarget.QualifiedModuleName)
                    .Where(d => d.IsMember() 
                                    || d.IsField() 
                                    || d.IsModuleConstant() 
                                    || d.DeclarationType.Equals(DeclarationType.UserDefinedType)
                                    || d.DeclarationType.Equals(DeclarationType.Enumeration))
                    .GroupBy(key => key.IdentifierName);

            var moveableMembers = new List<IMoveableMemberSet>();
            foreach (var group in groupsByIdentifier)
            {
                var newMoveable = new MoveableMemberSet(group.ToList());
                newMoveable.IsSelected = newMoveable.IdentifierName == moveTarget.IdentifierName;

                var idRefs = new List<IdentifierReference>();
                foreach (var member in newMoveable.Members.Where(m => m.IsMember()))
                {
                    idRefs = FindDirectTypeReferences(member).ToList();

                    var memberContainedReferences = _declarationFinderProvider.DeclarationFinder.IdentifierReferences(member.QualifiedName)
                        .Where(rf => !(rf.Declaration.DeclarationType.HasFlag(DeclarationType.Parameter) || rf.Declaration == rf.ParentScoping));
                    idRefs.AddRange(memberContainedReferences);
                }

                newMoveable.DirectReferences = idRefs;

                moveableMembers.Add(newMoveable);
            }

            var constants = moveableMembers.Where(m => m.Member.IsModuleConstant()).ToList();
            foreach (var moveableMember in constants)
            {
                var lExprContexts = moveableMember.Member.Context.GetDescendents<VBAParser.LExprContext>();
                if (lExprContexts.Any())
                {
                    var otherConstantIdentifierRefs = constants.Where(c => c != moveableMember)
                                                        .SelectMany(oc => oc.Member.References);

                    moveableMember.DirectReferences = otherConstantIdentifierRefs
                                    .Where(rf => lExprContexts.Contains(rf.Context.Parent));
                }
            }

            foreach (var moveableMember in moveableMembers.Where(m => m.Member.IsField()).ToList())
            {
                moveableMember.DirectReferences = FindDirectTypeReferences(moveableMember.Member).ToList();
            }
            return moveableMembers;
        }

        private IEnumerable<IdentifierReference> FindDirectTypeReferences(Declaration member)
        {
            var types = _declarationFinderProvider.DeclarationFinder.Members(member.QualifiedModuleName)
                .Where(m => m.DeclarationType.Equals(DeclarationType.UserDefinedType) || m.DeclarationType.Equals(DeclarationType.Enumeration));

            var idRefs = new List<IdentifierReference>();
            foreach (var typeReference in types.AllReferences())
            {
                if (member.AsTypeDeclaration?.Equals(typeReference.Declaration) ?? false)
                {
                    var memberAsTypeContext = member.Context.GetDescendent<VBAParser.AsTypeClauseContext>();
                    var referenceAsTypeContext = typeReference.Context.GetAncestor<VBAParser.AsTypeClauseContext>();
                    if (memberAsTypeContext.Equals(referenceAsTypeContext))
                    {
                        idRefs.Add(typeReference);
                    }
                }
            }
            return idRefs;
        }

        public IMovedContentProvider CreateMovedContentProvider()
        {
            return new MovedContentProvider();
        }

    }
}
