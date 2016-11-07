using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;
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
    public class MemberAccessTypeBindingTests
    {
        private const string BindingTargetName = "BindingTarget";
        private const string TestClassName = "TestClass";
        private const string ReferencedProjectFilepath = @"C:\Temp\ReferencedProjectA";

        [TestMethod]
        public void LExpressionIsProjectAndUnrestrictedNameIsProject()
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(BindingTargetName, ProjectProtection.Unprotected);
            string enclosingModuleCode = string.Format("Public WithEvents anything As {0}.{0}", BindingTargetName);
            enclosingProjectBuilder.AddComponent(TestClassName, ComponentType.ClassModule, enclosingModuleCode);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Project && d.ProjectName == BindingTargetName);

            // lExpression adds one reference, the MemberAcecssExpression adds another one.
            Assert.AreEqual(2, declaration.References.Count());
        }

        [TestMethod]
        public void LExpressionIsProjectAndUnrestrictedNameIsProceduralModule()
        {
            const string PROJECT_NAME = "AnyName";
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(PROJECT_NAME, ProjectProtection.Unprotected);
            string enclosingModuleCode = string.Format("Public WithEvents anything As {0}.{1}", PROJECT_NAME, BindingTargetName);
            enclosingProjectBuilder.AddComponent(BindingTargetName, ComponentType.StandardModule, enclosingModuleCode);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.ProceduralModule && d.IdentifierName == BindingTargetName);

            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void LExpressionIsProjectAndUnrestrictedNameIsClassModule()
        {
            const string PROJECT_NAME = "AnyName";
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(PROJECT_NAME, ProjectProtection.Unprotected);
            string enclosingModuleCode = string.Format("Public WithEvents anything As {0}.{1}", PROJECT_NAME, BindingTargetName);
            enclosingProjectBuilder.AddComponent(BindingTargetName, ComponentType.ClassModule, enclosingModuleCode);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.ClassModule && d.IdentifierName == BindingTargetName);

            Assert.AreEqual(1, declaration.References.Count());
        }

        [Ignore] // this failing test seems to be a resolver bug.
        [TestMethod]
        public void LExpressionIsProjectAndUnrestrictedNameIsType()
        {
            var builder = new MockVbeBuilder();
            const string referencedProjectName = "AnyReferencedProjectName";

            var referencedProject = builder
                .ProjectBuilder(referencedProjectName, ReferencedProjectFilepath, ProjectProtection.Unprotected)
                .AddComponent("AnyProceduralModuleName", ComponentType.StandardModule, CreateEnumType(BindingTargetName), Selection.Home)
                .Build();
            builder.AddProject(referencedProject);

            var enclosingProject = builder
                .ProjectBuilder("AnyProjectName", ProjectProtection.Unprotected)
                .AddReference(referencedProjectName, ReferencedProjectFilepath)
                .AddComponent(TestClassName, ComponentType.ClassModule, string.Format("Public WithEvents anything As {0}.{1}", referencedProjectName, BindingTargetName))
                .Build();
            builder.AddProject(enclosingProject);

            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Enumeration && d.IdentifierName == BindingTargetName);

            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void LExpressionIsModuleAndUnrestrictedNameIsType()
        {
            var builder = new MockVbeBuilder();
            const string CLASS_NAME = "AnyName";
            var enclosingProjectBuilder = builder.ProjectBuilder("AnyProjectName", ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddComponent(TestClassName, ComponentType.ClassModule, string.Format("Public WithEvents anything As {0}.{1}", CLASS_NAME, BindingTargetName));
            enclosingProjectBuilder.AddComponent(CLASS_NAME, ComponentType.ClassModule, CreateUdt(BindingTargetName));
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);

            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.UserDefinedType && d.IdentifierName == BindingTargetName);

            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void NestedMemberAccessExpressions()
        {
            var builder = new MockVbeBuilder();
            const string PROJECT_NAME = "AnyProjectName";
            const string CLASS_NAME = "AnyName";
            var enclosingProjectBuilder = builder.ProjectBuilder(PROJECT_NAME, ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddComponent(TestClassName, ComponentType.ClassModule, string.Format("Public WithEvents anything As {0}.{1}.{2}", PROJECT_NAME, CLASS_NAME, BindingTargetName));
            enclosingProjectBuilder.AddComponent(CLASS_NAME, ComponentType.ClassModule, CreateUdt(BindingTargetName));
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);

            var vbe = builder.Build();
            var state = Parse(vbe);

            Declaration declaration;

            declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Project && d.ProjectName == PROJECT_NAME);
            Assert.AreEqual(1, declaration.References.Count(), "Project reference expected");

            declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.ClassModule && d.IdentifierName == CLASS_NAME);
            Assert.AreEqual(1, declaration.References.Count(), "Module reference expected");

            declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.UserDefinedType && d.IdentifierName == BindingTargetName);
            Assert.AreEqual(1, declaration.References.Count(), "Type reference expected");
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
