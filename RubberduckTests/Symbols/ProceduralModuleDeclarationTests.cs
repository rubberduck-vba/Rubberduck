using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Symbols
{
    [TestClass]
    public class ProceduralModuleDeclarationTests
    {
        [TestMethod]
        public void ProceduralModulesHaveDeclarationTypeProceduralModule()
        {
            var projectDeclaration = GetTestProject("testProject");
            var proceduralModule = GetTestProceduralModule(projectDeclaration, "testModule", true, null);

            Assert.IsTrue(proceduralModule.DeclarationType.HasFlag(DeclarationType.ProceduralModule));
        }

            private static ProjectDeclaration GetTestProject(string name)
            {
                var qualifiedProjectName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new ProjectDeclaration(qualifiedProjectName, name, true, null);
            }

                private static QualifiedModuleName StubQualifiedModuleName()
                {
                    return new QualifiedModuleName("dummy", "dummy", "dummy");
                }

            private static ProceduralModuleDeclaration GetTestProceduralModule(Declaration projectDeclatation, string name, bool isUserDefined, Attributes attributes)
            {
                var qualifiedProceduralModuleMemberName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new ProceduralModuleDeclaration(qualifiedProceduralModuleMemberName, projectDeclatation, name, isUserDefined, null, attributes);
            }


        [TestMethod]
        public void ByDefaultProceduralModulesAreNotPrivate()
        {
            var projectDeclaration = GetTestProject("testProject");
            var proceduralModule = GetTestProceduralModule(projectDeclaration, "testModule", true, null);

            Assert.IsFalse(proceduralModule.IsPrivateModule);
        }

    }
}
