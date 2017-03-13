using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.VBEditor;

namespace RubberduckTests.Refactoring.ExtractMethod
{
    [TestClass]
    public class ExtractMethodRuleInSelectionTests
    {
        [TestClass]
        public class WhenSetValidFlagIsCalledWhenTheReferenceIsUsedAfter : ExtractMethodRuleInSelectionTests
        {
            [TestMethod]
            [TestCategory("ExtractMethodRuleTests")]
            public void shouldNotSetFlag()
            {
                var usedSelection = new Selection(4, 1, 7, 10);
                var referenceSelection = new Selection(8, 1, 8, 10);
                IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null);

                var SUT = new ExtractMethodRuleInSelection();
                var flag = SUT.setValidFlag(reference, usedSelection);

                Assert.AreEqual(0, flag);

            }
        }

        [TestClass]
        public class WhenSetValidFlagIsCalledWhenTheReferenceIsUsedBefore : ExtractMethodRuleInSelectionTests
        {
            [TestMethod]
            [TestCategory("ExtractMethodRuleTests")]
            public void shouldNotSetFlag()
            {
                var usedSelection = new Selection(4, 1, 7, 10);
                var referenceSelection = new Selection(3, 1, 3, 10);
                IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null);

                var SUT = new ExtractMethodRuleInSelection();
                var flag = SUT.setValidFlag(reference, usedSelection);

                Assert.AreEqual(0, flag);

            }
        }

    }
}
