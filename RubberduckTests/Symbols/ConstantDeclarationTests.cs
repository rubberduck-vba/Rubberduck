using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace RubberduckTests.Symbols
{
    [TestFixture]
    public class ConstantDeclarationTests
    {
        [Test]
        [Category("Resolver")]
        public void ExpressionReturnsTheContructorInjectedValue()
        {
            var value = "testtest";
            var constantName =  new QualifiedMemberName(StubQualifiedModuleName(),"testConstant");
            var constantDeclaration = new ConstantDeclaration(constantName, null, "test", "test", null, "test", null, Accessibility.Implicit, DeclarationType.Constant, value, null, Selection.Home, true);

            Assert.AreEqual(value, constantDeclaration.Expression);
        }
            
            private static QualifiedModuleName StubQualifiedModuleName()
            {
                return new QualifiedModuleName("dummy", "dummy", "dummy");
            }
    }
}
