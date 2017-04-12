using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;
using System;
using System.Linq;
using System.Threading;
using Moq;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Binding
{
    [TestClass]
    public class SimpleNameDefaultBindingTests
    {
        private const string BindingTargetName = "BindingTarget";
        private const string TestClassName = "TestClass";
        private const string ReferencedProjectFilepath = @"C:\Temp\ReferencedProjectA";

        [TestMethod]
        public void EnclosingProcedureComesBeforeEnclosingModule()
        {
            string testCode = string.Format(@"
Public Sub Test()
    Dim {0} As Long
    Dim a As String * {0}
End Sub", BindingTargetName);

            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(BindingTargetName, ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddComponent(TestClassName, ComponentType.ClassModule, testCode);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            var state = Parse(vbe);
            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Variable && d.IdentifierName == BindingTargetName);
            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void EnclosingModuleComesBeforeEnclosingProject()
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(BindingTargetName, ProjectProtection.Unprotected);
            string code = CreateEnumType(BindingTargetName) + Environment.NewLine + CreateTestProcedure(BindingTargetName);
            enclosingProjectBuilder.AddComponent(TestClassName, ComponentType.ClassModule, code);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            var state = Parse(vbe);
            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Enumeration && d.IdentifierName == BindingTargetName);
            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void EnclosingProjectComesBeforeOtherModuleInEnclosingProject()
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(BindingTargetName, ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddComponent(TestClassName, ComponentType.ClassModule, CreateTestProcedure(BindingTargetName));
            enclosingProjectBuilder.AddComponent("AnyModule", ComponentType.StandardModule, CreateEnumType(BindingTargetName));
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            var state = Parse(vbe);
            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Project && d.IdentifierName == BindingTargetName);

            Assert.AreEqual(state.Status, ParserState.Ready);
            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void OtherModuleInEnclosingProjectComesBeforeReferencedProjectModule()
        {
            var builder = new MockVbeBuilder();
            const string referencedProjectName = "AnyReferencedProjectName";

            var referencedProjectBuilder = builder.ProjectBuilder(referencedProjectName, ReferencedProjectFilepath, ProjectProtection.Unprotected);
            referencedProjectBuilder.AddComponent(BindingTargetName, ComponentType.ClassModule, string.Empty);
            var referencedProject = referencedProjectBuilder.Build();
            builder.AddProject(referencedProject);

            var enclosingProjectBuilder = builder.ProjectBuilder("AnyProjectName", ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddReference(referencedProjectName, ReferencedProjectFilepath);
            enclosingProjectBuilder.AddComponent(TestClassName, ComponentType.ClassModule, CreateTestProcedure(BindingTargetName));
            enclosingProjectBuilder.AddComponent("AnyModule", ComponentType.StandardModule, CreateEnumType(BindingTargetName));
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);

            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Enumeration && d.IdentifierName == BindingTargetName);

            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void ReferencedProjectModuleComesBeforeReferencedProjectType()
        {
            var builder = new MockVbeBuilder();
            const string referencedProjectName = "AnyReferencedProjectName";

            var referencedProject = builder
                .ProjectBuilder(referencedProjectName, ReferencedProjectFilepath, ProjectProtection.Unprotected)
                .AddComponent(BindingTargetName, ComponentType.StandardModule, CreateEnumType(BindingTargetName))
                .Build();
            builder.AddProject(referencedProject);

            var enclosingProject = builder
                .ProjectBuilder("AnyProjectName", ProjectProtection.Unprotected)
                .AddReference(referencedProjectName, ReferencedProjectFilepath)
                .AddComponent(TestClassName, ComponentType.ClassModule, CreateTestProcedure(BindingTargetName))
                .Build();
            builder.AddProject(enclosingProject);

            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.ProceduralModule && d.IdentifierName == BindingTargetName);

            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void ReferencedProjectClassNotMarkedAsGlobalClassModuleIsNotReferenced()
        {
            var builder = new MockVbeBuilder();
            const string referencedProjectName = "AnyReferencedProjectName";

            var referencedProjectBuilder = builder.ProjectBuilder(referencedProjectName, ReferencedProjectFilepath, ProjectProtection.Unprotected);
            referencedProjectBuilder.AddComponent("AnyName", ComponentType.ClassModule, CreateEnumType(BindingTargetName));
            var referencedProject = referencedProjectBuilder.Build();
            builder.AddProject(referencedProject);

            var enclosingProjectBuilder = builder.ProjectBuilder("AnyProjectName", ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddReference(referencedProjectName, ReferencedProjectFilepath);
            enclosingProjectBuilder.AddComponent(TestClassName, ComponentType.ClassModule, CreateTestProcedure(BindingTargetName));
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);

            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Enumeration && d.IdentifierName == BindingTargetName);

            Assert.AreEqual(0, declaration.References.Count());
        }

        private static RubberduckParserState Parse(Mock<IVBE> vbe)
        {
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));
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
    }
}
