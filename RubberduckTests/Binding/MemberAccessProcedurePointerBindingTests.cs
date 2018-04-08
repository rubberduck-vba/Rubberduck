using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;
using System;
using System.Linq;
using Moq;
using System.Threading;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Binding
{
    [TestFixture]
    public class MemberAccessProcedurePointerBindingTests
    {
        private const string BINDING_TARGET_NAME = "BindingTarget";
        private const string TEST_CLASS_NAME = "TestClass";

        [Category("Binding")]
        [Test]
        public void ProceduralModuleWithAccessibleMember()
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(BINDING_TARGET_NAME, ProjectProtection.Unprotected);
            string code = CreateCaller(TEST_CLASS_NAME) + Environment.NewLine + CreateCallee();
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, ComponentType.StandardModule, code);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {
                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.ProceduralModule && d.IdentifierName == TEST_CLASS_NAME);
                Assert.AreEqual(1, declaration.References.Count(), "Procedural Module should have reference");

                declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Function && d.IdentifierName == BINDING_TARGET_NAME);
                Assert.AreEqual(1, declaration.References.Count(), "Function should have reference");
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

        private string CreateCaller(string moduleName)
        {
            return $@"
Declare PtrSafe Function EnumWindows Lib ""user32"" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long

Public Sub Caller()
    EnumWindows AddressOf {moduleName}.{BINDING_TARGET_NAME}, 1
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
