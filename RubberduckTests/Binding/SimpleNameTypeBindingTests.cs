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
    public class SimpleNameTypeBindingTests
    {
        private const string BINDING_TARGET_NAME = "BindingTarget";
        private const string TEST_CLASS_NAME = "TestClass";
        private static readonly string ReferencedProjectFilepath = string.Empty; // must be an empty string

        [TestMethod]
        public void EnclosingModuleComesBeforeEnclosingProject()
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(BINDING_TARGET_NAME, ProjectProtection.Unprotected);
            string enclosingModuleCode = "Public WithEvents anything As " + BINDING_TARGET_NAME + Environment.NewLine + CreateEnumType(BINDING_TARGET_NAME);
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.ClassModule, enclosingModuleCode);
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
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.ClassModule, "Public WithEvents anything As  " + BINDING_TARGET_NAME);
            enclosingProjectBuilder.AddComponent("AnyModule", ComponentType.StandardModule, CreateEnumType(BINDING_TARGET_NAME));
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Project && d.IdentifierName == BINDING_TARGET_NAME);

            Assert.AreEqual(state.Status, ParserState.Ready);
            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void OtherModuleInEnclosingProjectComesBeforeReferencedProjectModule()
        {
            var builder = new MockVbeBuilder();
            const string REFERENCED_PROJECT_NAME = "AnyReferencedProjectName";

            var referencedProjectBuilder = builder.ProjectBuilder(REFERENCED_PROJECT_NAME, ReferencedProjectFilepath, ProjectProtection.Unprotected);
            referencedProjectBuilder.AddComponent(BINDING_TARGET_NAME, ComponentType.ClassModule, string.Empty);
            var referencedProject = referencedProjectBuilder.Build();
            builder.AddProject(referencedProject);

            var enclosingProjectBuilder = builder.ProjectBuilder("AnyProjectName", ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddReference(REFERENCED_PROJECT_NAME, ReferencedProjectFilepath);
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.ClassModule, "Public WithEvents anything As " + BINDING_TARGET_NAME);
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

            var referencedProjectBuilder = builder.ProjectBuilder(REFERENCED_PROJECT_NAME, ReferencedProjectFilepath, ProjectProtection.Unprotected);
            referencedProjectBuilder.AddComponent(BINDING_TARGET_NAME, ComponentType.StandardModule, CreateEnumType(BINDING_TARGET_NAME));
            var referencedProject = referencedProjectBuilder.Build();
            builder.AddProject(referencedProject);

            var enclosingProjectBuilder = builder.ProjectBuilder("AnyProjectName", ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddReference(REFERENCED_PROJECT_NAME, ReferencedProjectFilepath);
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.ClassModule, "Public WithEvents anything As " + BINDING_TARGET_NAME);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);

            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.ProceduralModule && d.IdentifierName == BINDING_TARGET_NAME);

            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void ReferencedProjectType()
        {
            var builder = new MockVbeBuilder();
            const string referencedProjectName = "Referenced";

            var referencedCode = CreateEnumType(BINDING_TARGET_NAME);
            const string enclosingCode = "Public AnyEnum As " + BINDING_TARGET_NAME;

            var referencedProjectBuilder = builder.ProjectBuilder(referencedProjectName, ReferencedProjectFilepath, ProjectProtection.Unprotected);
            referencedProjectBuilder.AddComponent("AnyName", ComponentType.StandardModule, referencedCode);
            var referencedProject = referencedProjectBuilder.Build();
            builder.AddProject(referencedProject);

            var enclosingProjectBuilder = builder.ProjectBuilder("Enclosing", ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddReference(referencedProjectName, ReferencedProjectFilepath);
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.ClassModule, enclosingCode);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);

            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Enumeration && d.IdentifierName == BINDING_TARGET_NAME);

            Assert.AreEqual(1, declaration.References.Count());
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
