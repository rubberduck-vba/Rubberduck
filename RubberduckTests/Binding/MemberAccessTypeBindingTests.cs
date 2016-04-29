using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;
using System.Linq;

namespace RubberduckTests.Binding
{
    [TestClass]
    public class MemberAccessTypeBindingTests
    {
        private const string BINDING_TARGET_NAME = "BindingTarget";
        private const string TEST_CLASS_NAME = "TestClass";
        private const string REFERENCED_PROJECT_FILEPATH = @"C:\Temp\ReferencedProjectA";

        [TestClass]
        public class ResolverTests
        {
            [TestMethod]
            public void LExpressionIsProjectAndUnrestrictedNameIsProject()
            {
                var builder = new MockVbeBuilder();
                var enclosingProjectBuilder = builder.ProjectBuilder(BINDING_TARGET_NAME, vbext_ProjectProtection.vbext_pp_none);
                string enclosingModuleCode = string.Format("Implements {0}.{0}", BINDING_TARGET_NAME);
                enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, vbext_ComponentType.vbext_ct_ClassModule, enclosingModuleCode);
                var enclosingProject = enclosingProjectBuilder.Build();
                builder.AddProject(enclosingProject);
                var vbe = builder.Build();
                var state = Parse(vbe);

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Project && d.Project.Name == BINDING_TARGET_NAME);

                // lExpression adds one reference, the MemberAcecssExpression adds another one.
                Assert.AreEqual(1, declaration.References.Count());
            }

            [TestMethod]
            public void LExpressionIsProjectAndUnrestrictedNameIsProceduralModule()
            {
                const string PROJECT_NAME = "AnyName";
                var builder = new MockVbeBuilder();
                var enclosingProjectBuilder = builder.ProjectBuilder(PROJECT_NAME, vbext_ProjectProtection.vbext_pp_none);
                string enclosingModuleCode = string.Format("Implements {0}.{1}", PROJECT_NAME, BINDING_TARGET_NAME);
                enclosingProjectBuilder.AddComponent(BINDING_TARGET_NAME, vbext_ComponentType.vbext_ct_StdModule, enclosingModuleCode);
                var enclosingProject = enclosingProjectBuilder.Build();
                builder.AddProject(enclosingProject);
                var vbe = builder.Build();
                var state = Parse(vbe);

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.ProceduralModule && d.IdentifierName == BINDING_TARGET_NAME);

                Assert.AreEqual(1, declaration.References.Count());
            }

            [TestMethod]
            public void LExpressionIsProjectAndUnrestrictedNameIsClassModule()
            {
                const string PROJECT_NAME = "AnyName";
                var builder = new MockVbeBuilder();
                var enclosingProjectBuilder = builder.ProjectBuilder(PROJECT_NAME, vbext_ProjectProtection.vbext_pp_none);
                string enclosingModuleCode = string.Format("Implements {0}.{1}", PROJECT_NAME, BINDING_TARGET_NAME);
                enclosingProjectBuilder.AddComponent(BINDING_TARGET_NAME, vbext_ComponentType.vbext_ct_ClassModule, enclosingModuleCode);
                var enclosingProject = enclosingProjectBuilder.Build();
                builder.AddProject(enclosingProject);
                var vbe = builder.Build();
                var state = Parse(vbe);

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.ClassModule && d.IdentifierName == BINDING_TARGET_NAME);

                Assert.AreEqual(1, declaration.References.Count());
            }

            [TestMethod]
            public void LExpressionIsProjectAndUnrestrictedNameIsType()
            {
                var builder = new MockVbeBuilder();
                const string REFERENCED_PROJECT_NAME = "AnyReferencedProjectName";

                var referencedProjectBuilder = builder.ProjectBuilder(REFERENCED_PROJECT_NAME, REFERENCED_PROJECT_FILEPATH, vbext_ProjectProtection.vbext_pp_none);
                referencedProjectBuilder.AddComponent("AnyProceduralModuleName", vbext_ComponentType.vbext_ct_StdModule, CreateEnumType(BINDING_TARGET_NAME));
                var referencedProject = referencedProjectBuilder.Build();
                builder.AddProject(referencedProject);

                var enclosingProjectBuilder = builder.ProjectBuilder("AnyProjectName", vbext_ProjectProtection.vbext_pp_none);
                enclosingProjectBuilder.AddReference(REFERENCED_PROJECT_NAME, REFERENCED_PROJECT_FILEPATH);
                enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, vbext_ComponentType.vbext_ct_ClassModule, string.Format("Implements {0}.{1}", REFERENCED_PROJECT_NAME, BINDING_TARGET_NAME));
                var enclosingProject = enclosingProjectBuilder.Build();
                builder.AddProject(enclosingProject);

                var vbe = builder.Build();
                var state = Parse(vbe);

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Enumeration && d.IdentifierName == BINDING_TARGET_NAME);

                Assert.AreEqual(1, declaration.References.Count());
            }

            [TestMethod]
            public void LExpressionIsModuleAndUnrestrictedNameIsType()
            {
                var builder = new MockVbeBuilder();
                const string CLASS_NAME = "AnyName";
                var enclosingProjectBuilder = builder.ProjectBuilder("AnyProjectName", vbext_ProjectProtection.vbext_pp_none);
                enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, vbext_ComponentType.vbext_ct_ClassModule, string.Format("Implements {0}.{1}", CLASS_NAME, BINDING_TARGET_NAME));
                enclosingProjectBuilder.AddComponent(CLASS_NAME, vbext_ComponentType.vbext_ct_ClassModule, CreateUdt(BINDING_TARGET_NAME));
                var enclosingProject = enclosingProjectBuilder.Build();
                builder.AddProject(enclosingProject);

                var vbe = builder.Build();
                var state = Parse(vbe);

                var declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.UserDefinedType && d.IdentifierName == BINDING_TARGET_NAME);

                Assert.AreEqual(1, declaration.References.Count());
            }

            [TestMethod]
            public void NestedMemberAccessExpressions()
            {
                var builder = new MockVbeBuilder();
                const string PROJECT_NAME = "AnyProjectName";
                const string CLASS_NAME = "AnyName";
                var enclosingProjectBuilder = builder.ProjectBuilder(PROJECT_NAME, vbext_ProjectProtection.vbext_pp_none);
                enclosingProjectBuilder.AddComponent(TEST_CLASS_NAME, vbext_ComponentType.vbext_ct_ClassModule, string.Format("Implements {0}.{1}.{2}", PROJECT_NAME, CLASS_NAME, BINDING_TARGET_NAME));
                enclosingProjectBuilder.AddComponent(CLASS_NAME, vbext_ComponentType.vbext_ct_ClassModule, CreateUdt(BINDING_TARGET_NAME));
                var enclosingProject = enclosingProjectBuilder.Build();
                builder.AddProject(enclosingProject);

                var vbe = builder.Build();
                var state = Parse(vbe);

                Declaration declaration;

                declaration  = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.Project && d.Project.Name == PROJECT_NAME);
                Assert.AreEqual(1, declaration.References.Count(), "Project reference expected");

                declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.ClassModule && d.IdentifierName == CLASS_NAME);
                Assert.AreEqual(1, declaration.References.Count(), "Module reference expected");

                declaration = state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.UserDefinedType && d.IdentifierName == BINDING_TARGET_NAME);
                Assert.AreEqual(1, declaration.References.Count(), "Type reference expected");
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
