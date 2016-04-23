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
    public class SimpleNameTypeBindingTests
    {
        private const string BINDING_TARGET_NAME = "BindingTarget";
        private const string TEST_CLASS_NAME = "TestClass";
        private const string REFERENCED_PROJECT_FILEPATH = @"C:\Temp\ReferencedProjectA";        

        [TestClass]
        public class ResolverTests
        {
            [TestMethod]
            public void EnclosingModuleComesBeforeEnclosingProject()
            {
                var builder = new MockVbeBuilder();
                var enclosingProjectBuilder = builder.ProjectBuilder(BINDING_TARGET_NAME, vbext_ProjectProtection.vbext_pp_none);
                string enclosingModuleCode = "Implements " + BINDING_TARGET_NAME + Environment.NewLine + CreateEnumType(BINDING_TARGET_NAME);
                enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, vbext_ComponentType.vbext_ct_ClassModule, enclosingModuleCode);
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
                var enclosingProjectBuilder = builder.ProjectBuilder(BINDING_TARGET_NAME, vbext_ProjectProtection.vbext_pp_none);
                enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, vbext_ComponentType.vbext_ct_ClassModule, "Implements " + BINDING_TARGET_NAME);
                enclosingProjectBuilder.AddComponent("AnyModule", vbext_ComponentType.vbext_ct_StdModule, CreateEnumType(BINDING_TARGET_NAME));
                var enclosingProject = enclosingProjectBuilder.Build();
                builder.AddProject(enclosingProject);
                var vbe = builder.Build();
                var state = Parse(vbe);

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Project && d.IdentifierName == BINDING_TARGET_NAME);

                Assert.AreEqual(1, declaration.References.Count());
            }

            [TestMethod]
            public void OtherModuleInEnclosingProjectComesBeforeReferencedProjectModule()
            {
                var builder = new MockVbeBuilder();
                const string REFERENCED_PROJECT_NAME = "AnyReferencedProjectName";

                var referencedProjectBuilder = builder.ProjectBuilder(REFERENCED_PROJECT_NAME, REFERENCED_PROJECT_FILEPATH, vbext_ProjectProtection.vbext_pp_none);
                referencedProjectBuilder.AddComponent(BINDING_TARGET_NAME, vbext_ComponentType.vbext_ct_ClassModule, string.Empty);
                var referencedProject = referencedProjectBuilder.Build();
                builder.AddProject(referencedProject);

                var enclosingProjectBuilder = builder.ProjectBuilder("AnyProjectName", vbext_ProjectProtection.vbext_pp_none);
                enclosingProjectBuilder.AddReference(REFERENCED_PROJECT_NAME, REFERENCED_PROJECT_FILEPATH);
                enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, vbext_ComponentType.vbext_ct_ClassModule, "Implements " + BINDING_TARGET_NAME);
                enclosingProjectBuilder.AddComponent("AnyModule", vbext_ComponentType.vbext_ct_StdModule, CreateEnumType(BINDING_TARGET_NAME));
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

                var referencedProjectBuilder = builder.ProjectBuilder(REFERENCED_PROJECT_NAME, REFERENCED_PROJECT_FILEPATH, vbext_ProjectProtection.vbext_pp_none);
                referencedProjectBuilder.AddComponent(BINDING_TARGET_NAME, vbext_ComponentType.vbext_ct_StdModule, CreateEnumType(BINDING_TARGET_NAME));
                var referencedProject = referencedProjectBuilder.Build();
                builder.AddProject(referencedProject);

                var enclosingProjectBuilder = builder.ProjectBuilder("AnyProjectName", vbext_ProjectProtection.vbext_pp_none);
                enclosingProjectBuilder.AddReference(REFERENCED_PROJECT_NAME, REFERENCED_PROJECT_FILEPATH);
                enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, vbext_ComponentType.vbext_ct_ClassModule, "Implements " + BINDING_TARGET_NAME);
                var enclosingProject = enclosingProjectBuilder.Build();
                builder.AddProject(enclosingProject);

                var vbe = builder.Build();
                var state = Parse(vbe);

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.ProceduralModule && d.IdentifierName == BINDING_TARGET_NAME);

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
}
