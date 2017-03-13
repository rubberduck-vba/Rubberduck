using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.VBEditor;

namespace RubberduckTests.Refactoring.ExtractMethod
{
    [TestClass]
    public class ExtractMethodRuleExternalReferenceTests
    {

        [TestClass]
        public class WhenSetValidFlagIsCalledWhenTheReferenceIsInternal : ExtractMethodRuleExternalReferenceTests
        {
            [TestMethod]
            [TestCategory("ExtractMethodRuleTests")]
            public void shouldSetTheFlag()
            {
                var usedSelection = new Selection(4, 1, 7, 10);
                var referenceSelection = new Selection(8, 1, 8, 10);

                var decQualifiedMemberName = new QualifiedMemberName(new QualifiedModuleName(), "");
                var decSelection = new Selection(5, 1, 5, 10);
                var referenceDeclaration = new Declaration(decQualifiedMemberName, null, "", "", "",
                    false, false, Accessibility.Friend, DeclarationType.ClassModule,
                    context: null, selection: decSelection, isArray: false, asTypeContext: null);

                IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, referenceDeclaration);

                var SUT = new ExtractMethodRuleExternalReference();
                var flag = SUT.setValidFlag(reference, usedSelection);
                var expected = ((byte)ExtractMethodRuleFlags.IsExternallyReferenced);
                Assert.AreEqual(expected, flag);
            }
        }

        [TestClass]
        public class WhenSetValidFlagIsCalledWhenTheReferenceIsUsedInternallyOnly : ExtractMethodRuleExternalReferenceTests
        {
            [TestMethod]
            [TestCategory("ExtractMethodRuleTests")]
            public void shouldNotSetFlag()
            {
                var usedSelection = new Selection(4, 1, 7, 10);

                var referenceSelection = new Selection(7, 1, 7, 10);

                var decQualifiedMemberName = new QualifiedMemberName(new QualifiedModuleName(), "");
                var decSelection = new Selection(5, 1, 5, 10);
                var referenceDeclaration = new Declaration(decQualifiedMemberName, null, "", "", "",
                    false, false, Accessibility.Friend, DeclarationType.ClassModule,
                    context: null, selection: decSelection, isArray: false, asTypeContext: null);

                IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, referenceDeclaration);

                var SUT = new ExtractMethodRuleExternalReference();
                var flag = SUT.setValidFlag(reference, usedSelection);

                Assert.AreEqual(0, flag);

            }
        }

    }
}
