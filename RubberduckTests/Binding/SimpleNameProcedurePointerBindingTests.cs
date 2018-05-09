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
    public class SimpleNameProcedurePointerBindingTests
    {
        private const string BINDING_TARGET_NAME = "BindingTarget";
        private const string TEST_CLASS_NAME = "TestClass";

        [Category("Binding")]
        [Test]
        public void EnclosingModuleComesBeforeEnclosingProject()
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(BINDING_TARGET_NAME, ProjectProtection.Unprotected);
            string code = CreateCaller() + Environment.NewLine + CreateCallee();
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.ClassModule, code);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Function && d.IdentifierName == BINDING_TARGET_NAME);

                Assert.AreEqual(1, declaration.References.Count());
            }
        }

        [Category("Binding")]
        [Test]
        public void EnclosingProjectComesBeforeOtherProceduralModule()
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(BINDING_TARGET_NAME, ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.ClassModule, CreateCaller());
            enclosingProjectBuilder.AddComponent("AnyProceduralModule", ComponentType.StandardModule, CreateCallee());
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Project && d.IdentifierName == BINDING_TARGET_NAME);

                Assert.AreEqual(state.Status, ParserState.Ready);
                Assert.AreEqual(1, declaration.References.Count());
            }
        }

        [Category("Binding")]
        [Test]
        public void OtherProceduralModule()
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder("AnyProjectName", ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.ClassModule, CreateCaller());
            enclosingProjectBuilder.AddComponent("AnyProceduralModule", ComponentType.StandardModule, CreateCallee());
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Function && d.IdentifierName == BINDING_TARGET_NAME);

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

        private string CreateCaller()
        {
            return $@"
Declare PtrSafe Function EnumWindows Lib ""user32"" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long

Public Sub Caller()
    EnumWindows AddressOf {BINDING_TARGET_NAME}, 1
End Sub
";
        }

        private string CreateCallee()
        {
            return $@"
Public Function {BINDING_TARGET_NAME}() As LongPtr
End Function
";
        }
    }
}
