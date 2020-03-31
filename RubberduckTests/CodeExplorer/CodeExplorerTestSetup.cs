using System;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.ComManagement;
using RubberduckTests.AddRemoveReferences;

namespace RubberduckTests.CodeExplorer
{
    internal static class CodeExplorerTestSetup
    {
        private static readonly List<Declaration> ProjectOne;
        private static readonly List<Declaration> ProjectTwo;

        public static readonly IProjectsProvider ProjectOneProvider;
        public static readonly IProjectsProvider ProjectTwoProvider;

        public const string TestProjectOneName = "TestProject1";
        public const string TestProjectTwoName = "TestProject2";
        public const string TestDocumentName = "TestDocument1";
        public const string TestModuleName = "TestModule1";
        public const string TestClassName = "TestClass1";
        public const string TestUserFormName = "TestUserForm1";

        public static List<Declaration> TestProjectOneDeclarations => new List<Declaration>(ProjectOne);

        public static List<Declaration> TestProjectTwoDeclarations => new List<Declaration>(ProjectTwo);

        public static List<Declaration> TestComponentDeclarations(this List<Declaration> declarations, string moduleName)
        {
            var projectDeclaration = declarations.Single(declaration => declaration.DeclarationType == DeclarationType.Project);
            var candidates = CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration, ref declarations);

            return candidates.Where(declaration =>
                declaration.DeclarationType != DeclarationType.Project &&
                declaration.QualifiedModuleName.ComponentName.Equals(moduleName)).ToList();
        }

        public static List<Declaration> TestProjectWithComponentDeclarations(this List<Declaration> declarations, IEnumerable<string> moduleNames, out Declaration projectDeclaration)
        {
            projectDeclaration = declarations.Single(declaration => declaration.DeclarationType == DeclarationType.Project);
            var candidates = CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration, ref declarations).ToList();

            var output = new List<Declaration> { projectDeclaration };

            foreach (var component in moduleNames)
            {
                output.AddRange(candidates.Where(declaration =>
                    declaration.DeclarationType != DeclarationType.Project &&
                    declaration.QualifiedModuleName.ComponentName.Equals(component)));
            }
            return output;
        }

        public static List<Declaration> TestProjectWithComponentRemoved(this List<Declaration> declarations, string moduleName)
        {
            var projectDeclaration = declarations.Single(declaration => declaration.DeclarationType == DeclarationType.Project);
            var removing = declarations.ToList().TestComponentDeclarations(moduleName);
            var candidates = CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration, ref declarations).ToList();

            return candidates.Except(removing).ToList();
        }

        public static List<Declaration> TestMemberDeclarations(this List<Declaration> declarations, 
            string memberName, out Declaration memberDeclaration, DeclarationType type = DeclarationType.Member)
        {
            var member = declarations.Single(declaration =>
                declaration.IdentifierName.Equals(memberName) && 
                (type == DeclarationType.Member || declaration.DeclarationType == type));

            var projectDeclaration = declarations.Single(declaration => declaration.DeclarationType == DeclarationType.Project);

            memberDeclaration = member;

            var candidates = CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration, ref declarations);

            return candidates.Where(declaration =>
                ReferenceEquals(declaration, member) ||
                ReferenceEquals(declaration.ParentDeclaration, member)).ToList();
        }

        public static List<Declaration> TestProjectWithMemberRemoved(this List<Declaration> declarations, 
            string memberName, out Declaration componentDeclaration, DeclarationType type = DeclarationType.Member)
        {
            var projectDeclaration = declarations.Single(declaration => declaration.DeclarationType == DeclarationType.Project);
            var removing = declarations.ToList().TestMemberDeclarations(memberName, out var member, type);

            componentDeclaration = member.ParentDeclaration;

            var candidates = CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration, ref declarations).ToList();

            return candidates.Except(removing).ToList();
        }

        public static List<Declaration> TestSubMemberDeclarations(this List<Declaration> declarations,
            string subMemberName, out Declaration subMemberDeclaration)
        {
            subMemberDeclaration = declarations.Single(declaration => declaration.IdentifierName.Equals(subMemberName));
            return new List<Declaration> { subMemberDeclaration };
        }

        private static readonly Dictionary<string, string> CodeByModuleName = new Dictionary<string, string>
        {
            { TestDocumentName, CodeExplorerTestCode.TestDocumentCode },
            { TestModuleName, CodeExplorerTestCode.TestModuleCode },
            { TestClassName, CodeExplorerTestCode.TestClassCode },
            { TestUserFormName, CodeExplorerTestCode.TestUserFormCode }
        };

        private static readonly Dictionary<string, ComponentType> ComponentTypeByModuleName = new Dictionary<string, ComponentType>
        {
            { TestDocumentName, ComponentType.Document },
            { TestModuleName, ComponentType.StandardModule },
            { TestClassName, ComponentType.ClassModule },
            { TestUserFormName, ComponentType.UserForm }
        };

        public static List<Declaration> TestProjectWithFolderStructure(IEnumerable<(string Name, string Folder)> modules, out Declaration projectDeclaration, out RubberduckParserState state)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder(TestProjectOneName, ProjectProtection.Unprotected);

            foreach (var (name, folder) in modules)
            {
                var code = string.IsNullOrEmpty(folder)
                    ? CodeByModuleName[name]
                    : string.Join(Environment.NewLine, $"'@Folder({folder.ToVbaStringLiteral()})", CodeByModuleName[name]);

                var type = ComponentTypeByModuleName[name];
                if (type == ComponentType.UserForm)
                {
                    project.MockUserFormBuilder(TestUserFormName, code).AddFormToProjectBuilder();
                    continue;
                }

                project.AddComponent(name, type, code);
            }

            builder.AddProject(project.Build());
            state = MockParser.CreateAndParse(builder.Build().Object);
            var output = state.AllUserDeclarations.ToList();

            projectDeclaration = output.Single(declaration => declaration.DeclarationType == DeclarationType.Project);
            return output;
        }

        static CodeExplorerTestSetup()
        {
            var builder = new MockVbeBuilder();

            var project = builder.ProjectBuilder(TestProjectOneName, ProjectProtection.Unprotected)
                .AddComponent(TestDocumentName, ComponentType.Document, CodeExplorerTestCode.TestDocumentCode)
                .AddComponent(TestModuleName, ComponentType.StandardModule, CodeExplorerTestCode.TestModuleCode)
                .AddComponent(TestClassName, ComponentType.ClassModule, CodeExplorerTestCode.TestClassCode);

            project.MockUserFormBuilder(TestUserFormName, CodeExplorerTestCode.TestUserFormCode).AddFormToProjectBuilder();
            builder.AddProject(project.Build());

            var state = MockParser.CreateAndParse(builder.Build().Object);
            ProjectOne = state.AllUserDeclarations.ToList();
            ProjectOneProvider = state.ProjectsProvider;

            builder = new MockVbeBuilder();

            project = builder.ProjectBuilder(TestProjectTwoName, ProjectProtection.Unprotected)
                .AddComponent(TestDocumentName, ComponentType.Document, CodeExplorerTestCode.TestDocumentCode)
                .AddComponent(TestModuleName, ComponentType.StandardModule, CodeExplorerTestCode.TestModuleCode)
                .AddComponent(TestClassName, ComponentType.ClassModule, CodeExplorerTestCode.TestClassCode);

            project.MockUserFormBuilder(TestUserFormName, CodeExplorerTestCode.TestUserFormCode).AddFormToProjectBuilder();
            builder.AddProject(project.Build());

            state = MockParser.CreateAndParse(builder.Build().Object);
            ProjectTwo = state.AllUserDeclarations.ToList();
            ProjectTwoProvider = state.ProjectsProvider;
        }

        public static List<Declaration> GetProjectDeclarationsWithReferences(bool libraries, bool projects, out RubberduckParserState state)
        {
            var builder = new MockVbeBuilder();

            var project = builder.ProjectBuilder(TestProjectOneName, ProjectProtection.Unprotected)
                .AddComponent(TestDocumentName, ComponentType.Document, CodeExplorerTestCode.TestDocumentCode);

            if (libraries)
            {
                foreach (var library in AddRemoveReferencesSetup.DummyReferencesList)
                {
                    project.AddReference(library.Name, library.FullPath, library.Major, library.Minor);
                }
            }

            if (projects)
            {
                foreach (var reference in AddRemoveReferencesSetup.DummyProjectsList)
                {
                    project.AddReference(reference.Name, reference.FullPath, reference.Major, reference.Minor, false, ReferenceKind.Project);
                }
            }

            builder.AddProject(project.Build());
            state = MockParser.CreateAndParse(builder.Build().Object);
            return state.AllUserDeclarations.ToList();
        }

        public static List<Declaration> GetAllChildDeclarations(this ICodeExplorerNode node)
        {
            var output = new List<Declaration>();

            // The project declaration is the fallback for non-declaration nodes, so exclude it from those nodes.
            if (node is CodeExplorerProjectViewModel ||
                node.Declaration.DeclarationType != DeclarationType.Project)
            {
                output.Add(node.Declaration);
            }

            foreach (var child in node.Children)
            {
                output.AddRange(child.GetAllChildDeclarations());
            }

            return output;
        }

        public static Declaration ShallowCopy(this Declaration declaration)
        {
            return new Declaration(
                declaration.QualifiedName,
                declaration.ParentDeclaration,
                declaration.ParentScopeDeclaration,
                declaration.AsTypeName,
                declaration.TypeHint,
                declaration.IsSelfAssigned,
                declaration.IsWithEvents,
                declaration.Accessibility,
                declaration.DeclarationType,
                declaration.Context,
                declaration.AttributesPassContext,
                declaration.Selection,
                declaration.IsArray,
                declaration.AsTypeContext,
                declaration.IsUserDefined,
                declaration.Annotations,
                declaration.Attributes,
                declaration.IsUndeclared);
        }
    }
}
