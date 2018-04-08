using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.VBEditor;

namespace RubberduckTests.Refactoring.ExtractMethod
{

    [TestFixture]
    public class ExtractMethodRuleIsAssignedInSelectionTests
    {
        [TestFixture]
        public class WhenSetValidFlagIsCalledWhenTheReferenceIsInSelection : ExtractMethodRuleIsAssignedInSelectionTests
        {
            [TestFixture]
            public class AndIsAssigned : WhenSetValidFlagIsCalledWhenTheReferenceIsInSelection
            {
                [Test]
                [Category("ExtractMethodRuleTests")]
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

            [TestFixture]
            public class AndIsNotAssigned : WhenSetValidFlagIsCalledWhenTheReferenceIsInSelection
            {
                [Test]
                [Category("ExtractMethodRuleTests")]
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

        [TestFixture]
        public class WhenSetValidFlagIsCalledWhenTheReferenceIsAssigned : ExtractMethodRuleIsAssignedInSelectionTests
        {

            [TestFixture]
            public class AndIsBeforeSelection : WhenSetValidFlagIsCalledWhenTheReferenceIsAssigned
            {
                [Test]
                [Category("ExtractMethodRuleTests")]
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

            [TestFixture]
            public class AndIsAfterSelection : WhenSetValidFlagIsCalledWhenTheReferenceIsAssigned
            {
                [Test]
                [Category("ExtractMethodRuleTests")]
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
