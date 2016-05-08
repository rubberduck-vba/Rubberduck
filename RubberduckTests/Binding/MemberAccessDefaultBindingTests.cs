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
    public class MemberAccessDefaultBindingTests
    {
        private const string BINDING_TARGET_LEXPRESSION = "BindingTarget";
        private const string BINDING_TARGET_UNRESTRICTEDNAME = "UnrestrictedName";
        private const string TEST_CLASS_NAME = "TestClass";
        private const string REFERENCED_PROJECT_FILEPATH = @"C:\Temp\ReferencedProjectA";

        [TestClass]
        public class ResolverTests
        {
            [TestMethod]
            public void LExpressionIsVariablePropertyOrFunction()
            {
                var builder = new MockVbeBuilder();
                var enclosingProjectBuilder = builder.ProjectBuilder("Any Project", vbext_ProjectProtection.vbext_pp_none);
                string code = string.Format("Public Sub Test() {0} Dim {1} As {2} {0} Call {1}.{3} {0}End Sub", Environment.NewLine, "AnyName", BINDING_TARGET_LEXPRESSION, BINDING_TARGET_UNRESTRICTEDNAME);
                enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, vbext_ComponentType.vbext_ct_ClassModule, code);
                enclosingProjectBuilder.AddComponent(BINDING_TARGET_LEXPRESSION, vbext_ComponentType.vbext_ct_ClassModule, CreateFunction(BINDING_TARGET_UNRESTRICTEDNAME));
                var enclosingProject = enclosingProjectBuilder.Build();
                builder.AddProject(enclosingProject);
                var vbe = builder.Build();
                var state = Parse(vbe);

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Function && d.IdentifierName == BINDING_TARGET_UNRESTRICTEDNAME);

                Assert.AreEqual(1, declaration.References.Count());
            }

            [TestMethod]
            public void LExpressionIsProject()
            {
                const string PROJECT_NAME = "AnyProject";
                var builder = new MockVbeBuilder();
                var enclosingProjectBuilder = builder.ProjectBuilder(PROJECT_NAME, vbext_ProjectProtection.vbext_pp_none);
                string code = string.Format("Public Sub Test() {0} Call {1}.{2} {0}End Sub", Environment.NewLine, PROJECT_NAME, BINDING_TARGET_UNRESTRICTEDNAME);
                enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, vbext_ComponentType.vbext_ct_ClassModule, code);
                enclosingProjectBuilder.AddComponent("Anymodulename", vbext_ComponentType.vbext_ct_StdModule, CreateFunction(BINDING_TARGET_UNRESTRICTEDNAME));
                var enclosingProject = enclosingProjectBuilder.Build();
                builder.AddProject(enclosingProject);
                var vbe = builder.Build();
                var state = Parse(vbe);

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Function && d.IdentifierName == BINDING_TARGET_UNRESTRICTEDNAME);

                Assert.AreEqual(1, declaration.References.Count());
            }

            [TestMethod]
            public void LExpressionIsProceduralModule()
            {
                const string PROJECT_NAME = "AnyProject";
                var builder = new MockVbeBuilder();
                var enclosingProjectBuilder = builder.ProjectBuilder(PROJECT_NAME, vbext_ProjectProtection.vbext_pp_none);
                string code = string.Format("Public Sub Test() {0} Call {1}.{2} {0}End Sub", Environment.NewLine, BINDING_TARGET_LEXPRESSION, BINDING_TARGET_UNRESTRICTEDNAME);
                enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, vbext_ComponentType.vbext_ct_ClassModule, code);
                enclosingProjectBuilder.AddComponent(BINDING_TARGET_LEXPRESSION, vbext_ComponentType.vbext_ct_StdModule, CreateFunction(BINDING_TARGET_UNRESTRICTEDNAME));
                var enclosingProject = enclosingProjectBuilder.Build();
                builder.AddProject(enclosingProject);
                var vbe = builder.Build();
                var state = Parse(vbe);

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Function && d.IdentifierName == BINDING_TARGET_UNRESTRICTEDNAME);

                Assert.AreEqual(1, declaration.References.Count());
            }

            [TestMethod]
            public void LExpressionIsEnum()
            {
                const string PROJECT_NAME = "AnyProject";
                var builder = new MockVbeBuilder();
                var enclosingProjectBuilder = builder.ProjectBuilder(PROJECT_NAME, vbext_ProjectProtection.vbext_pp_none);
                string code = string.Format("Public Sub Test() {0} a = {1}.{2} {0}End Sub", Environment.NewLine, BINDING_TARGET_LEXPRESSION, BINDING_TARGET_UNRESTRICTEDNAME);
                enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, vbext_ComponentType.vbext_ct_ClassModule, code);
                enclosingProjectBuilder.AddComponent("AnyModule", vbext_ComponentType.vbext_ct_StdModule, CreateEnumType(BINDING_TARGET_LEXPRESSION, BINDING_TARGET_UNRESTRICTEDNAME));
                var enclosingProject = enclosingProjectBuilder.Build();
                builder.AddProject(enclosingProject);
                var vbe = builder.Build();
                var state = Parse(vbe);

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.EnumerationMember && d.IdentifierName == BINDING_TARGET_UNRESTRICTEDNAME);

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

            private string CreateFunction(string functionName)
            {
                return string.Format(@"
Public Function {0}() As Variant
    TestEnumMember
End Function
", functionName);
            }

            private string CreateEnumType(string typeName, string memberName)
            {
                return string.Format(@"
Public Enum {0}
    {1}
End Enum
", typeName, memberName);
            }
        }
    }
}
