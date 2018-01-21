using NUnit.Framework;
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
    [TestFixture]
    public class MemberAccessDefaultBindingTests
    {
        private const string BINDING_TARGET_LEXPRESSION = "BindingTarget";
        private const string BINDING_TARGET_UNRESTRICTEDNAME = "UnrestrictedName";
        private const string TEST_CLASS_NAME = "TestClass";
        private const string REFERENCED_PROJECT_FILEPATH = @"C:\Temp\ReferencedProjectA";

        [Category("Binding")]
        [Test]
        public void LExpressionIsVariablePropertyOrFunction()
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder("Any Project", ProjectProtection.Unprotected);
            string code = string.Format("Public Sub Test() {0} Dim {1} As {2} {0} Call {1}.{3} {0}End Sub", Environment.NewLine, "AnyName", BINDING_TARGET_LEXPRESSION, BINDING_TARGET_UNRESTRICTEDNAME);
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.ClassModule, code);
            enclosingProjectBuilder.AddComponent(BINDING_TARGET_LEXPRESSION, ComponentType.ClassModule, CreateFunction(BINDING_TARGET_UNRESTRICTEDNAME));
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Function && d.IdentifierName == BINDING_TARGET_UNRESTRICTEDNAME);

                Assert.AreEqual(1, declaration.References.Count());
            }
        }

        [Category("Binding")]
        [Test]
        public void LExpressionIsProject()
        {
            const string PROJECT_NAME = "AnyProject";
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(PROJECT_NAME, ProjectProtection.Unprotected);
            string code = string.Format("Public Sub Test() {0} Call {1}.{2} {0}End Sub", Environment.NewLine, PROJECT_NAME, BINDING_TARGET_UNRESTRICTEDNAME);
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.ClassModule, code);
            enclosingProjectBuilder.AddComponent("Anymodulename", ComponentType.StandardModule, CreateFunction(BINDING_TARGET_UNRESTRICTEDNAME));
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {
                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Function && d.IdentifierName == BINDING_TARGET_UNRESTRICTEDNAME);

                Assert.AreEqual(1, declaration.References.Count());
            }
        }

        [Category("Binding")]
        [Test]
        public void LExpressionIsProceduralModule()
        {
            const string PROJECT_NAME = "AnyProject";
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(PROJECT_NAME, ProjectProtection.Unprotected);
            string code = string.Format("Public Sub Test() {0} Call {1}.{2} {0}End Sub", Environment.NewLine, BINDING_TARGET_LEXPRESSION, BINDING_TARGET_UNRESTRICTEDNAME);
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.ClassModule, code);
            enclosingProjectBuilder.AddComponent(BINDING_TARGET_LEXPRESSION, ComponentType.StandardModule, CreateFunction(BINDING_TARGET_UNRESTRICTEDNAME));
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {
                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Function && d.IdentifierName == BINDING_TARGET_UNRESTRICTEDNAME);

                Assert.AreEqual(1, declaration.References.Count());
            }
        }

        [Category("Binding")]
        [Test]
        public void LExpressionIsEnum()
        {
            const string PROJECT_NAME = "AnyProject";
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(PROJECT_NAME, ProjectProtection.Unprotected);
            string code = string.Format("Public Sub Test() {0} a = {1}.{2} {0}End Sub", Environment.NewLine, BINDING_TARGET_LEXPRESSION, BINDING_TARGET_UNRESTRICTEDNAME);
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.ClassModule, code);
            enclosingProjectBuilder.AddComponent("AnyModule", ComponentType.StandardModule, CreateEnumType(BINDING_TARGET_LEXPRESSION, BINDING_TARGET_UNRESTRICTEDNAME));
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {
                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.EnumerationMember && d.IdentifierName == BINDING_TARGET_UNRESTRICTEDNAME);

                Assert.AreEqual(1, declaration.References.Count());
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

        private string CreateFunction(string functionName)
        {
            return $@"
Public Function {functionName}() As Variant
    TestEnumMember
End Function
";
        }

        private string CreateEnumType(string typeName, string memberName)
        {
            return $@"
Public Enum {typeName}
    {memberName}
End Enum
";
        }
    }
}
