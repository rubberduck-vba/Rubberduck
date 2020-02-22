using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Symbols
{
    [TestFixture]
    public class ProceduralModuleDeclarationTests
    {
        [Test]
        [Category("Resolver")]
        public void ProceduralModulesHaveDeclarationTypeProceduralModule()
        {
            var projectDeclaration = GetTestProject("testProject");
            var proceduralModule = GetTestProceduralModule(projectDeclaration, "testModule", true, null);

            Assert.IsTrue(proceduralModule.DeclarationType.HasFlag(DeclarationType.ProceduralModule));
        }

            private static ProjectDeclaration GetTestProject(string name)
            {
                var qualifiedProjectName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new ProjectDeclaration(qualifiedProjectName, name, true);
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


        [Test]
        [Category("Resolver")]
        public void ByDefaultProceduralModulesAreNotPrivate()
        {
            var projectDeclaration = GetTestProject("testProject");
            var proceduralModule = GetTestProceduralModule(projectDeclaration, "testModule", true, null);

            Assert.IsFalse(proceduralModule.IsPrivateModule);
        }

    }
}
