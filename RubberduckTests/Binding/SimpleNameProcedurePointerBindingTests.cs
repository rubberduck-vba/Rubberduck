using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;
using System;
using System.Linq;

namespace RubberduckTests.Binding
{
    [TestClass]
    public class SimpleNameProcedurePointerBindingTests
    {
        private const string BINDING_TARGET_NAME = "BindingTarget";
        private const string TEST_CLASS_NAME = "TestClass";

        [TestMethod]
        public void EnclosingModuleComesBeforeEnclosingProject()
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(BINDING_TARGET_NAME, vbext_ProjectProtection.vbext_pp_none);
            string code = CreateCaller() + Environment.NewLine + CreateCallee();
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, vbext_ComponentType.vbext_ct_ClassModule, code);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Function && d.IdentifierName == BINDING_TARGET_NAME);

            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void EnclosingProjectComesBeforeOtherProceduralModule()
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(BINDING_TARGET_NAME, vbext_ProjectProtection.vbext_pp_none);
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, vbext_ComponentType.vbext_ct_ClassModule, CreateCaller());
            enclosingProjectBuilder.AddComponent("AnyProceduralModule", vbext_ComponentType.vbext_ct_StdModule, CreateCallee());
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Project && d.IdentifierName == BINDING_TARGET_NAME);

            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void OtherProceduralModule()
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder("AnyProjectName", vbext_ProjectProtection.vbext_pp_none);
            enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, vbext_ComponentType.vbext_ct_ClassModule, CreateCaller());
            enclosingProjectBuilder.AddComponent("AnyProceduralModule", vbext_ComponentType.vbext_ct_StdModule, CreateCallee());
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Function && d.IdentifierName == BINDING_TARGET_NAME);

            Assert.AreEqual(1, declaration.References.Count());
        }

        private static RubberduckParserState Parse(Moq.Mock<VBE> vbe)
        {
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());
            parser.Parse();
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser state should be 'Ready', but returns '{0}'.", parser.State.Status);
            }
            var state = parser.State;
            return state;
        }

        private string CreateCaller()
        {
            return string.Format(@"
Declare PtrSafe Function EnumWindows Lib ""user32"" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long

Public Sub Caller()
    EnumWindows AddressOf {0}, 1
End Sub
", BINDING_TARGET_NAME);
        }

        private string CreateCallee()
        {
            return string.Format(@"
Public Function {0}() As LongPtr
End Function
", BINDING_TARGET_NAME);
        }
    }
}
