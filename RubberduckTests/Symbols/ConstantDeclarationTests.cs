using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace RubberduckTests.Symbols
{
    [TestClass]
    public class ConstantDeclarationTests
    {
        [TestMethod]
        public void ExpressionReturnsTheContructorInjectedValue()
        {
            var value = "testtest";
            var constantName =  new QualifiedMemberName(StubQualifiedModuleName(),"testConstant");
            var constantDeclaration = new ConstantDeclaration(constantName, null, "test", "test", null, "test", null, Accessibility.Implicit, DeclarationType.Constant, value, null, Selection.Home, true);

            Assert.AreEqual<string>(value, constantDeclaration.Expression);
        }
            
            private static QualifiedModuleName StubQualifiedModuleName()
            {
                return new QualifiedModuleName("dummy", "dummy", "dummy");
            }
    }
}
