using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.CodeExplorer
{
    internal static class CodeExplorerTestSetup
    {
        private static readonly List<Declaration> ProjectOne;
        private static readonly List<Declaration> ProjectTwo;

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
            var candidates = CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration, declarations);

            return candidates.Where(declaration =>
                declaration.DeclarationType != DeclarationType.Project &&
                declaration.QualifiedModuleName.ComponentName.Equals(moduleName)).ToList();
        }

        public static List<Declaration> TestMemberDeclarations(this List<Declaration> declarations, 
            string memberName, out Declaration memberDeclaration, DeclarationType type = DeclarationType.Member)
        {
            var member = declarations.Single(declaration =>
                declaration.IdentifierName.Equals(memberName) && 
                (type == DeclarationType.Member || declaration.DeclarationType == type));

            var projectDeclaration = declarations.Single(declaration => declaration.DeclarationType == DeclarationType.Project);

            memberDeclaration = member;

            var candidates = CodeExplorerProjectViewModel.ExtractTrackedDeclarationsForProject(projectDeclaration, declarations);

            return candidates.Where(declaration =>
                ReferenceEquals(declaration, member) ||
                ReferenceEquals(declaration.ParentDeclaration, member)).ToList();
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

            var parser = MockParser.CreateAndParse(builder.Build().Object);
            ProjectOne = parser.AllUserDeclarations.ToList();

            builder = new MockVbeBuilder();

            project = builder.ProjectBuilder(TestProjectTwoName, ProjectProtection.Unprotected)
                .AddComponent(TestDocumentName, ComponentType.Document, CodeExplorerTestCode.TestDocumentCode)
                .AddComponent(TestModuleName, ComponentType.StandardModule, CodeExplorerTestCode.TestModuleCode)
                .AddComponent(TestClassName, ComponentType.ClassModule, CodeExplorerTestCode.TestClassCode);

            project.MockUserFormBuilder(TestUserFormName, CodeExplorerTestCode.TestUserFormCode).AddFormToProjectBuilder();
            builder.AddProject(project.Build());

            parser = MockParser.CreateAndParse(builder.Build().Object);
            ProjectTwo = parser.AllUserDeclarations.ToList();
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
