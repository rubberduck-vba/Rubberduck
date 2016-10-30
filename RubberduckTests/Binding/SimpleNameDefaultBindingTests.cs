using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;
using System;
using System.Linq;
using System.Threading;
using Moq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Binding
{
    [TestClass]
    public class SimpleNameDefaultBindingTests
    {
        private const string BINDING_TARGET_NAME = "BindingTarget";
        private const string TEST_CLASS_NAME = "TestClass";
        private const string REFERENCED_PROJECT_FILEPATH = @"C:\Temp\ReferencedProjectA";

        [TestMethod]
        public void EnclosingProcedureComesBeforeEnclosingModule()
        {
            string testCode = string.Format(@"
Public Sub Test()
    Dim {0} As Long
    Dim a As String * {0}
End Sub", BINDING_TARGET_NAME);

            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(BINDING_TARGET_NAME, ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.ClassModule, testCode);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            var state = Parse(vbe);
            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Variable && d.IdentifierName == BINDING_TARGET_NAME);
            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void EnclosingModuleComesBeforeEnclosingProject()
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(BINDING_TARGET_NAME, ProjectProtection.Unprotected);
            string code = CreateEnumType(BINDING_TARGET_NAME) + Environment.NewLine + CreateTestProcedure(BINDING_TARGET_NAME);
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.ClassModule, code);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            var state = Parse(vbe);
            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Enumeration && d.IdentifierName == BINDING_TARGET_NAME);
            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void EnclosingProjectComesBeforeOtherModuleInEnclosingProject()
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(BINDING_TARGET_NAME, ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.ClassModule, CreateTestProcedure(BINDING_TARGET_NAME));
            enclosingProjectBuilder.AddComponent("AnyModule", ComponentType.StandardModule, CreateEnumType(BINDING_TARGET_NAME));
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            var state = Parse(vbe);
            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Project && d.IdentifierName == BINDING_TARGET_NAME);
            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void OtherModuleInEnclosingProjectComesBeforeReferencedProjectModule()
        {
            var builder = new MockVbeBuilder();
            const string REFERENCED_PROJECT_NAME = "AnyReferencedProjectName";

            var referencedProjectBuilder = builder.ProjectBuilder(REFERENCED_PROJECT_NAME, REFERENCED_PROJECT_FILEPATH, ProjectProtection.Unprotected);
            referencedProjectBuilder.AddComponent(BINDING_TARGET_NAME, ComponentType.ClassModule, string.Empty);
            var referencedProject = referencedProjectBuilder.Build();
            builder.AddProject(referencedProject);

            var enclosingProjectBuilder = builder.ProjectBuilder("AnyProjectName", ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddReference(REFERENCED_PROJECT_NAME, REFERENCED_PROJECT_FILEPATH);
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.ClassModule, CreateTestProcedure(BINDING_TARGET_NAME));
            enclosingProjectBuilder.AddComponent("AnyModule", ComponentType.StandardModule, CreateEnumType(BINDING_TARGET_NAME));
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);

            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Enumeration && d.IdentifierName == BINDING_TARGET_NAME);

            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void ReferencedProjectModuleComesBeforeReferencedProjectType()
        {
            var builder = new MockVbeBuilder();
            const string REFERENCED_PROJECT_NAME = "AnyReferencedProjectName";

            var referencedProjectBuilder = builder.ProjectBuilder(REFERENCED_PROJECT_NAME, REFERENCED_PROJECT_FILEPATH, ProjectProtection.Unprotected);
            referencedProjectBuilder.AddComponent(BINDING_TARGET_NAME, ComponentType.StandardModule, CreateEnumType(BINDING_TARGET_NAME));
            var referencedProject = referencedProjectBuilder.Build();
            builder.AddProject(referencedProject);

            var enclosingProjectBuilder = builder.ProjectBuilder("AnyProjectName", ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddReference(REFERENCED_PROJECT_NAME, REFERENCED_PROJECT_FILEPATH);
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.ClassModule, CreateTestProcedure(BINDING_TARGET_NAME));
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);

            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.ProceduralModule && d.IdentifierName == BINDING_TARGET_NAME);

            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void ReferencedProjectClassNotMarkedAsGlobalClassModuleIsNotReferenced()
        {
            var builder = new MockVbeBuilder();
            const string REFERENCED_PROJECT_NAME = "AnyReferencedProjectName";

            var referencedProjectBuilder = builder.ProjectBuilder(REFERENCED_PROJECT_NAME, REFERENCED_PROJECT_FILEPATH, ProjectProtection.Unprotected);
            referencedProjectBuilder.AddComponent("AnyName", ComponentType.ClassModule, CreateEnumType(BINDING_TARGET_NAME));
            var referencedProject = referencedProjectBuilder.Build();
            builder.AddProject(referencedProject);

            var enclosingProjectBuilder = builder.ProjectBuilder("AnyProjectName", ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddReference(REFERENCED_PROJECT_NAME, REFERENCED_PROJECT_FILEPATH);
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.ClassModule, CreateTestProcedure(BINDING_TARGET_NAME));
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);

            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Enumeration && d.IdentifierName == BINDING_TARGET_NAME);

            Assert.AreEqual(0, declaration.References.Count());
        }

        private static RubberduckParserState Parse(Mock<IVBE> vbe)
        {
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser state should be 'Ready', but returns '{0}'.", parser.State.Status);
            }
            var state = parser.State;
            return state;
        }

        private string CreateTestProcedure(string bindingTarget)
        {
            return string.Format(@"
Public Sub Test()
    Dim a As String * {0}
End Sub
", bindingTarget);
        }

        private string CreateEnumType(string typeName)
        {
            return string.Format(@"
Public Enum {0}
    TestEnumMember
End Enum
", typeName);
        }

        private string CreateUdt(string typeName)
        {
            return string.Format(@"
Public Type {0}
    TestTypeMember As String
End Type
", typeName);
        }
    }
}
