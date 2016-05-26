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
    public class IndexDefaultBindingTests
    {
        private const string BINDING_TARGET_LEXPRESSION = "BindingTarget";
        private const string BINDING_TARGET_UNRESTRICTEDNAME = "UnrestrictedName";
        private const string TEST_CLASS_NAME = "TestClass";
        private const string REFERENCED_PROJECT_FILEPATH = @"C:\Temp\ReferencedProjectA";

        [TestMethod]
        public void RecursiveDefaultMember()
        {
            string callerModule = @"
Public Property Get Ok() As Klasse1
End Property

Public Sub Test()
    Call Ok(1)
End Sub
";

            string middleman = @"
Public Property Get Test1() As Klasse2
End Property
";

            string defaultMemberClass = @"
Public Property Get Test2(a As Integer)
    Test2 = 2
End Property
";

            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder("Any Project", vbext_ProjectProtection.vbext_pp_none);
            enclosingProjectBuilder.AddComponent("AnyModule1", vbext_ComponentType.vbext_ct_StdModule, callerModule);
            enclosingProjectBuilder.AddComponent("AnyClass", vbext_ComponentType.vbext_ct_ClassModule, middleman);
            enclosingProjectBuilder.AddComponent("AnyClass2", vbext_ComponentType.vbext_ct_ClassModule, defaultMemberClass);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.PropertyGet && d.IdentifierName == "Test2");

            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void NormalPropertyFunctionSubroutine()
        {
            string callerModule = @"
Public Sub Test()
    Call Test1(1)
End Sub
";

            string property = @"
Public Property Get Test1(a As Integer) As Long
End Property
";

            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder("Any Project", vbext_ProjectProtection.vbext_pp_none);
            enclosingProjectBuilder.AddComponent("AnyModule1", vbext_ComponentType.vbext_ct_StdModule, callerModule);
            enclosingProjectBuilder.AddComponent("AnyClass", vbext_ComponentType.vbext_ct_StdModule, property);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            var state = Parse(vbe);

            var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.PropertyGet && d.IdentifierName == "Test1");

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
    }
}
