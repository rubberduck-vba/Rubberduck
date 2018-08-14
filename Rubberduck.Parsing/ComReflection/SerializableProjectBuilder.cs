using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Parsing.ComReflection
{
    public class SerializableProjectBuilder : ISerializableProjectBuilder
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        private DeclarationFinder _finder;

        public SerializableProjectBuilder(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }


        public SerializableProject SerializableProject(ProjectDeclaration projectDeclaration)
        {
            _finder = _declarationFinderProvider.DeclarationFinder;

            var serializableProject = new SerializableProject(projectDeclaration);

            var projectName = projectDeclaration.QualifiedModuleName;
            var projectLevelDeclarationsByParent = ProjectLevelDeclarations(projectName)
                .Where(declaration => declaration.ParentDeclaration != null)
                .GroupBy(declaration => declaration.ParentDeclaration)
                .ToDictionary();

            if(projectLevelDeclarationsByParent.TryGetValue(projectDeclaration, out var nonModuleProjectChildren))
            {
                foreach (var projectLevelDeclaration in nonModuleProjectChildren.OrderBy(child => child.DeclarationType).ThenBy(child => child.IdentifierName ?? string.Empty))
                {
                    serializableProject.AddDeclaration(new SerializableDeclarationTree(projectLevelDeclaration));
                }
            }

            foreach (var module in ProjectModules(projectName).OrderBy(qmn => qmn.ComponentType).ThenBy(qmn => qmn.ComponentName ?? string.Empty))
            {
                var serializableModule = SerializableModule(module, projectDeclaration, projectLevelDeclarationsByParent);
                serializableProject.AddDeclaration(serializableModule);
            }


            return serializableProject;
        }

        private IEnumerable<Declaration> ProjectLevelDeclarations(QualifiedModuleName projectName)
        {
            return _finder.Members(projectName);
        }

        private IEnumerable<QualifiedModuleName> ProjectModules(QualifiedModuleName projectName)
        {
            return _finder.AllModules.Where(qmn => qmn.ProjectId == projectName.ProjectId && !qmn.Equals(projectName));
        }

        private SerializableDeclarationTree SerializableModule(QualifiedModuleName module, ProjectDeclaration project, Dictionary<Declaration, List<Declaration>> projectLevelDeclarationsByParent)
        {
            var members = _finder.Members(module).ToList();
            var membersByParent = members.Where(declaration => declaration.ParentDeclaration != null)
                .GroupBy(declaration => declaration.ParentDeclaration)
                .ToDictionary();

            var moduleDeclaration = membersByParent[project].Single();
            var serializableModule = SerializableTree(moduleDeclaration, membersByParent);

            if (projectLevelDeclarationsByParent.TryGetValue(moduleDeclaration, out var memberDeclarationsOnProjectLevel))
            {
                var orderedMemberDeclarations = memberDeclarationsOnProjectLevel
                    .OrderBy(declaration => declaration.DeclarationType)
                    .ThenBy(declaration => declaration.IdentifierName ?? string.Empty);
                serializableModule.AddChildren(orderedMemberDeclarations);
            }

            return serializableModule;
        }

        private SerializableDeclarationTree SerializableTree(Declaration declaration,
            IDictionary<Declaration, List<Declaration>> declarationsByParent)
        {
            var serializableDeclaration = new SerializableDeclarationTree(declaration);
            var childTrees = ChildTrees(declaration, declarationsByParent).OrderBy(tree => tree.Node.DeclarationType).ThenBy(tree => tree.Node.IdentifierName ?? string.Empty);
            serializableDeclaration.AddChildTrees(childTrees);

            return serializableDeclaration;
        }

        private IEnumerable<SerializableDeclarationTree> ChildTrees(Declaration parentDeclaration,
            IDictionary<Declaration, List<Declaration>> declarationsByParent)
        {
            var childTrees = new List<SerializableDeclarationTree>();

            if (!declarationsByParent.TryGetValue(parentDeclaration, out var childDeclarations))
            {
                return childTrees;
            }

            foreach (var childDeclaration in childDeclarations)
            {
                childTrees.Add(SerializableTree(childDeclaration, declarationsByParent));
            }

            return childTrees;
        }
    }
}