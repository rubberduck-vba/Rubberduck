using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace RubberduckTests.Symbols
{
    [TestFixture]
    public class ParameterDeclarationTests
    {
        [Test]
        [Category("Resolver")]
        public void ParametersHaveDeclarationTypeParameter()
        {
            var paramter = GetTestParameter("testParam", false, false, false);

            Assert.IsTrue(paramter.DeclarationType.HasFlag(DeclarationType.Parameter));
        }

            private static ParameterDeclaration GetTestParameter(string name, bool isOptional, bool isByRef, bool isParamArray)
            {
                var qualifiedParameterName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new ParameterDeclaration(qualifiedParameterName, null, "test", null,"test", isOptional,isByRef, false, isParamArray);
            }

                private static QualifiedModuleName StubQualifiedModuleName()
                {
                    return new QualifiedModuleName("dummy", "dummy", "dummy");
                }


        [Test]
        [Category("Resolver")]
        public void ParametersHaveImpliciteAccessibility()
        {
            var paramter = GetTestParameter("testParam", false, false, false);

            Assert.IsTrue(paramter.Accessibility.HasFlag(Accessibility.Implicit));
        }


        [Test]
        [Category("Resolver")]
        public void IsParamArrayCanBeSetPublicly()
        {
            var paramter = GetTestParameter("testParam", false, false, false);
            paramter.IsParamArray = true;

            Assert.IsTrue(paramter.IsParamArray);
        }

    }
}
