using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.VBEditor;

namespace RubberduckTests.Refactoring.ExtractMethod
{

    [TestClass]
    public class ExtractMethodRuleIsAssignedInSelectionTests
    {
        [TestClass]
        public class WhenSetValidFlagIsCalledWhenTheReferenceIsInSelection : ExtractMethodRuleIsAssignedInSelectionTests
        {
            [TestClass]
            public class AndIsAssigned : WhenSetValidFlagIsCalledWhenTheReferenceIsInSelection
            {
                [TestMethod]
                [TestCategory("ExtractMethodRuleTests")]
                public void shouldSetFlagIsAssigned()
                {
                    var usedSelection = new Selection(4, 1, 7, 10);
                    var referenceSelection = new Selection(6, 1, 6, 10);
                    var isAssigned = true;
                    IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null, isAssigned);

                    var SUT = new ExtractMethodRuleIsAssignedInSelection();
                    var flag = SUT.setValidFlag(reference, usedSelection);

                    Assert.AreEqual((byte)ExtractMethodRuleFlags.IsAssigned, flag);

                }
            }

            [TestClass]
            public class AndIsNotAssigned : WhenSetValidFlagIsCalledWhenTheReferenceIsInSelection
            {
                [TestMethod]
                [TestCategory("ExtractMethodRuleTests")]
                public void shouldNotSetFlag()
                {
                    var usedSelection = new Selection(4, 1, 7, 10);
                    var referenceSelection = new Selection(6, 1, 6, 10);
                    var isAssigned = false;
                    IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null, isAssigned);

                    var SUT = new ExtractMethodRuleIsAssignedInSelection();
                    var flag = SUT.setValidFlag(reference, usedSelection);

                    Assert.AreEqual(0, flag);

                }
            }

        }

        [TestClass]
        public class WhenSetValidFlagIsCalledWhenTheReferenceIsAssigned : ExtractMethodRuleIsAssignedInSelectionTests
        {

            [TestClass]
            public class AndIsBeforeSelection : WhenSetValidFlagIsCalledWhenTheReferenceIsAssigned
            {
                [TestMethod]
                [TestCategory("ExtractMethodRuleTests")]
                public void shouldSetFlagIsAssigned()
                {
                    var usedSelection = new Selection(4, 1, 7, 10);
                    var referenceSelection = new Selection(3, 1, 3, 10);
                    var isAssigned = true;
                    IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null, isAssigned);

                    var SUT = new ExtractMethodRuleIsAssignedInSelection();
                    var flag = SUT.setValidFlag(reference, usedSelection);

                    Assert.AreEqual(0, flag);

                }
            }

            [TestClass]
            public class AndIsAfterSelection : WhenSetValidFlagIsCalledWhenTheReferenceIsAssigned
            {
                [TestMethod]
                [TestCategory("ExtractMethodRuleTests")]
                public void shouldNotSetFlag()
                {
                    var usedSelection = new Selection(4, 1, 7, 10);
                    var referenceSelection = new Selection(9, 1, 9, 10);
                    var isAssigned = true;
                    IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null, isAssigned);

                    var SUT = new ExtractMethodRuleIsAssignedInSelection();
                    var flag = SUT.setValidFlag(reference, usedSelection);

                    Assert.AreEqual(0, flag);

                }
            }

        }

    }
}
