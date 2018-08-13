using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

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
            foreach (var projectLevelDeclaration in ProjectLevelDeclarations(projectName))
            {
                serializableProject.AddDeclaration(new SerializableDeclarationTree(projectLevelDeclaration));
            }

            foreach (var module in ProjectModules(projectName.ProjectId))
            {
                var serializableModule = SerializableModule(module);
                serializableProject.AddDeclaration(serializableModule);
            }


            return serializableProject;
        }

        private IEnumerable<Declaration> ProjectLevelDeclarations(QualifiedModuleName projectName)
        {
            return _finder.Members(projectName);
        }

        private IEnumerable<QualifiedModuleName> ProjectModules(string projectId)
        {
            return _finder.AllModules.Where(qmn => qmn.ProjectId == projectId);
        }

        private SerializableDeclarationTree SerializableModule(QualifiedModuleName module)
        {
            var members = _finder.Members(module).ToList();
            var moduleDeclaration = members.Single(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Module));
            var membersByParent = members.GroupBy(declaration => declaration.ParentDeclaration).ToDictionary();
            var serializableModule = SerializableTree(moduleDeclaration, membersByParent);

            return serializableModule;
        }

        private SerializableDeclarationTree SerializableTree(Declaration declaration,
            IDictionary<Declaration, List<Declaration>> declarationsByParent)
        {
            var serializableDeclaration = new SerializableDeclarationTree(declaration);
            var childTrees = ChildTrees(declaration, declarationsByParent);
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