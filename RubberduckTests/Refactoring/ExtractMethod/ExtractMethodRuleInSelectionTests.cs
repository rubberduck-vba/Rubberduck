using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.VBEditor;

namespace RubberduckTests.Refactoring.ExtractMethod
{
    [TestFixture]
    public class ExtractMethodRuleInSelectionTests
    {
        [TestFixture]
        public class WhenSetValidFlagIsCalledWhenTheReferenceIsUsedAfter : ExtractMethodRuleInSelectionTests
        {
            [Test]
            [Category("ExtractMethodRuleTests")]
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

        [TestFixture]
        public class WhenSetValidFlagIsCalledWhenTheReferenceIsUsedBefore : ExtractMethodRuleInSelectionTests
        {
            [Test]
            [Category("ExtractMethodRuleTests")]
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
