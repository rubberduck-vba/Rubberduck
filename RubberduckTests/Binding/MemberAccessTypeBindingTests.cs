using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;
using System.Linq;
using System.Threading;
using Moq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Binding
{
    [TestFixture]
    public class MemberAccessTypeBindingTests
    {
        private const string BindingTargetName = "BindingTarget";
        private const string TestClassName = "TestClass";
        private const string ReferencedProjectFilepath = @"C:\Temp\ReferencedProjectA";

        [Category("Binding")]
        [Test]
        public void LExpressionIsProjectAndUnrestrictedNameIsProject()
        {
            var enclosingModuleCode = string.Format("Public WithEvents anything As {0}.{0}", BindingTargetName);

            var builder = new MockVbeBuilder();
            var enclosingProject = builder
                .ProjectBuilder(BindingTargetName, ProjectProtection.Unprotected)
                .AddComponent(TestClassName, ComponentType.ClassModule, enclosingModuleCode)
                .Build();
            builder.AddProject(enclosingProject);

            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Project && d.ProjectName == BindingTargetName);

                // lExpression adds one reference, the MemberAcecssExpression adds another one.
                Assert.AreEqual(2, declaration.References.Count());
            }
        }

        [Category("Binding")]
        [Test]
        public void LExpressionIsProjectAndUnrestrictedNameIsProceduralModule()
        {
            const string projectName = "AnyName";
            var enclosingModuleCode = $"Public WithEvents anything As {projectName}.{BindingTargetName}";

            var builder = new MockVbeBuilder();
            var enclosingProject = builder
                .ProjectBuilder(projectName, ProjectProtection.Unprotected)
                .AddComponent(BindingTargetName, ComponentType.StandardModule, enclosingModuleCode)
                .Build();
            builder.AddProject(enclosingProject);

            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.ProceduralModule && d.IdentifierName == BindingTargetName);

                Assert.AreEqual(1, declaration.References.Count());
            }
        }

        [Category("Binding")]
        [Test]
        public void LExpressionIsProjectAndUnrestrictedNameIsClassModule()
        {
            const string projectName = "AnyName";
            var enclosingModuleCode = $"Public WithEvents anything As {projectName}.{BindingTargetName}";

            var builder = new MockVbeBuilder();
            var enclosingProject = builder
                .ProjectBuilder(projectName, ProjectProtection.Unprotected)
                .AddComponent(BindingTargetName, ComponentType.ClassModule, enclosingModuleCode)
                .Build();
            builder.AddProject(enclosingProject);

            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.ClassModule && d.IdentifierName == BindingTargetName);

                Assert.AreEqual(1, declaration.References.Count());
            }
        }

        [Category("Binding")]
        [Test]
        public void LExpressionIsProjectAndUnrestrictedNameIsType()
        {
            var builder = new MockVbeBuilder();
            const string referencedProjectName = "AnyReferencedProjectName";
            var code = $"Public WithEvents anything As {referencedProjectName}.{BindingTargetName}";

            var referencedProject = builder
                .ProjectBuilder(referencedProjectName, ReferencedProjectFilepath, ProjectProtection.Unprotected)
                .AddComponent("AnyProceduralModuleName", ComponentType.StandardModule, CreateEnumType(BindingTargetName), Selection.Home)
                .Build();
            builder.AddProject(referencedProject);

            var enclosingProject = builder
                .ProjectBuilder("AnyProjectName", ProjectProtection.Unprotected)
                .AddReference(referencedProjectName, ReferencedProjectFilepath, 0, 0)
                .AddComponent(TestClassName, ComponentType.ClassModule, code)
                .Build();
            builder.AddProject(enclosingProject);

            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Enumeration && d.IdentifierName == BindingTargetName);

                Assert.AreEqual(1, declaration.References.Count());
            }
        }

        [Category("Binding")]
        [Test]
        public void LExpressionIsModuleAndUnrestrictedNameIsType()
        {
            var builder = new MockVbeBuilder();
            const string className = "AnyName";

            var enclosingProject = builder
                .ProjectBuilder("AnyProjectName", ProjectProtection.Unprotected)
                .AddComponent(TestClassName, ComponentType.ClassModule,
                    $"Public WithEvents anything As {className}.{BindingTargetName}")
                .AddComponent(className, ComponentType.ClassModule, CreateUdt(BindingTargetName))
                .Build();
            builder.AddProject(enclosingProject);

            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.UserDefinedType && d.IdentifierName == BindingTargetName);

                Assert.AreEqual(1, declaration.References.Count());
            }
        }

        [Category("Binding")]
        [Test]
        public void NestedMemberAccessExpressions()
        {
            const string projectName = "AnyProjectName";
            const string className = "AnyName";

            var builder = new MockVbeBuilder();
            var enclosingProject = builder
                .ProjectBuilder(projectName, ProjectProtection.Unprotected)
                .AddComponent(TestClassName, ComponentType.ClassModule,
                    $"Public WithEvents anything As {projectName}.{className}.{BindingTargetName}")
                .AddComponent(className, ComponentType.ClassModule, CreateUdt(BindingTargetName))
                .Build();
            builder.AddProject(enclosingProject);

            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {

                Assert.AreEqual(state.Status, ParserState.Ready);

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Project && d.ProjectName == projectName);
                Assert.AreEqual(1, declaration.References.Count(), "Project reference expected");

                declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.ClassModule && d.IdentifierName == className);
                Assert.AreEqual(1, declaration.References.Count(), "Module reference expected");

                declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.UserDefinedType && d.IdentifierName == BindingTargetName);
                Assert.AreEqual(1, declaration.References.Count(), "Type reference expected");
            }
        }

        private static RubberduckParserState Parse(Mock<IVBE> vbe)
        {
            var parser = MockParser.Create(vbe.Object);
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
            return $@"
Public Enum {typeName}
    TestEnumMember
End Enum
";
        }

        private string CreateUdt(string typeName)
        {
            return $@"
Public Type {typeName}
    TestTypeMember As String
End Type
";
        }
    }
}
