using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.ExtractMethod
{
    // https://github.com/rubberduck-vba/Rubberduck/wiki/Extract-Method-Refactoring-%3A-Workings---Determining-what-params-to-move
    [TestClass]
    public class ExtractMethodParameterClassificationTests
    {

        #region variableInternalAndOnlyUsedInternally
        string codeSnippet = @"
Option explicit
Public Sub CodeWithDeclaration()
    Dim x As Long
    Dim z As Long

    z = 1
    x = 1 + 2
    DoNothing x

End Sub

Public Sub DoNothing(n As Long)
End Sub
";
        #endregion

        [TestClass]
        public class WhenClassifyingDeclarations : ExtractMethodParameterClassificationTests
        {

            [TestMethod] 
            public void shouldUseEachRuleInRulesCollectionToCheckEachReference()
            {
                QualifiedModuleName qualifiedModuleName;
                var state = MockParser.ParseString(codeSnippet, out qualifiedModuleName);
                var declaration = state.AllUserDeclarations.Where(decl => decl.IdentifierName == "x").First();
                var selection = new Selection(5, 1, 7, 20);
                var qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var emRule = new Mock<IExtractMethodRule>();
                emRule.Setup(emr => emr.setValidFlag(It.IsAny<IdentifierReference>(), It.IsAny<Selection>())).Returns(2);
                var emRules = new List<IExtractMethodRule>() { emRule.Object, emRule.Object };
                var sut = new ExtractMethodParameterClassification(emRules);

                //Act
                sut.classifyDeclarations(qSelection, declaration);

                //Assert
                // 2 rules on 2 references = 4 validation checks
                var expectedToVerify = 4;
                emRule.Verify(emr => emr.setValidFlag(It.IsAny<IdentifierReference>(), selection),
                    Times.Exactly(expectedToVerify));

            }
        }

        [TestClass]
        public class WhenExtractingParameters : ExtractMethodParameterClassificationTests
        {

            [TestMethod]
            public void shouldIncludeByValParams()
            {

                QualifiedModuleName qualifiedModuleName;
                var state = MockParser.ParseString(codeSnippet, out qualifiedModuleName);
                var declaration = state.AllUserDeclarations.Where(decl => decl.IdentifierName == "x").First();

                // Above setup is headache from lack of ability to extract declarations simply.
                // exact declaration and qSelection are irrelevant for this test and could be mocked.

                var emByRefRule = new Mock<IExtractMethodRule>();
                emByRefRule.Setup(em => em.setValidFlag(It.IsAny<IdentifierReference>(), It.IsAny<Selection>())).Returns(14);
                var emRules = new List<IExtractMethodRule>() { emByRefRule.Object };

                var qSelection = new QualifiedSelection();
                var SUT = new ExtractMethodParameterClassification(emRules);
                //Act
                SUT.classifyDeclarations(qSelection, declaration);
                var extractedParameter = SUT.ExtractedParameters.First();
                Assert.IsTrue(SUT.ExtractedParameters.Count() > 0);

                //Assert
                Assert.AreEqual(extractedParameter.Passed, ExtractedParameter.PassedBy.ByVal);

            }

            [TestMethod]
            public void shouldIncludeByRefParams()
            {

                QualifiedModuleName qualifiedModuleName;
                var state = MockParser.ParseString(codeSnippet, out qualifiedModuleName);
                var declaration = state.AllUserDeclarations.Where(decl => decl.IdentifierName == "x").First();
                var selection = new Selection(5, 1, 7, 20);
                var qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                // Above setup is headache from lack of ability to extract declarations simply.
                // exact declaration and qSelection are irrelevant for this test and could be mocked.

                var emByRefRule = new Mock<IExtractMethodRule>();
                emByRefRule.Setup(em => em.setValidFlag(It.IsAny<IdentifierReference>(), It.IsAny<Selection>())).Returns(10);
                var emRules = new List<IExtractMethodRule>() { emByRefRule.Object };

                var SUT = new ExtractMethodParameterClassification(emRules);
                //Act
                SUT.classifyDeclarations(qSelection, declaration);
                var extractedParameter = SUT.ExtractedParameters.First();
                Assert.IsTrue(SUT.ExtractedParameters.Count() > 0);

                //Assert
                Assert.AreEqual(extractedParameter.Passed, ExtractedParameter.PassedBy.ByRef);

            }
        }
    }
 } 