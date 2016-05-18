using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
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
                byte flag = new byte();
                var usedSelection = new Selection(4, 1, 7, 10);
                var referenceSelection = new Selection(8, 1, 8, 10);
                IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null);

                var SUT = new ExtractMethodRuleInSelection();
                SUT.setValidFlag(ref flag, reference, usedSelection);

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
                byte flag = new byte();
                var usedSelection = new Selection(4, 1, 7, 10);
                var referenceSelection = new Selection(3, 1, 3, 10);
                IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null);

                var SUT = new ExtractMethodRuleInSelection();
                SUT.setValidFlag(ref flag, reference, usedSelection);

                Assert.AreEqual(0, flag);

            }

        }
        [TestClass]
        public class WhenSetValidFlagIsCalledWhenTheReferenceIsInSelection : ExtractMethodRuleInSelectionTests
        {
            [TestMethod]
            [TestCategory("ExtractMethodRuleTests")]
            public void shouldSetFlagInSelection()
            {
                byte flag = new byte();
                var usedSelection = new Selection(4, 1, 7, 10);
                var referenceSelection = new Selection(5, 1, 5, 10);
                IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null);

                var SUT = new ExtractMethodRuleInSelection();
                SUT.setValidFlag(ref flag, reference, usedSelection);

                Assert.AreEqual((byte)ExtractMethodRuleFlags.InSelection, flag);

            }

        }

    }

    [TestClass]
    public class ExtractMethodRuleUsedAfterTests
    {
        [TestClass]
        public class WhenSetValidFlagIsCalledWhenTheReferenceIsUsedAfter : ExtractMethodRuleUsedAfterTests
        {
            [TestMethod]
            [TestCategory("ExtractMethodRuleTests")]
            public void shouldNotSetFlagUsedAfter()
            {
                byte flag = new byte();
                var usedSelection = new Selection(4, 1, 7, 10);
                var referenceSelection = new Selection(8, 1, 8, 10);
                IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null);

                var SUT = new ExtractMethodRuleUsedAfter();
                SUT.setValidFlag(ref flag, reference, usedSelection);

                Assert.AreEqual((byte)ExtractMethodRuleFlags.UsedAfter, flag);

            }
        }
        [TestClass]
        public class WhenSetValidFlagIsCalledWhenTheReferenceIsUsedBefore : ExtractMethodRuleUsedAfterTests
        {
            [TestMethod]
            [TestCategory("ExtractMethodRuleTests")]
            public void shouldNotSetFlag()
            {
                byte flag = new byte();
                var usedSelection = new Selection(4, 1, 7, 10);
                var referenceSelection = new Selection(3, 1, 3, 10);
                IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null);

                var SUT = new ExtractMethodRuleUsedAfter();
                SUT.setValidFlag(ref flag, reference, usedSelection);

                Assert.AreEqual(0, flag);

            }

        }
        [TestClass]
        public class WhenSetValidFlagIsCalledWhenTheReferenceIsInSelection : ExtractMethodRuleUsedAfterTests
        {
            [TestMethod]
            [TestCategory("ExtractMethodRuleTests")]
            public void shouldNotSetFlag()
            {
                byte flag = new byte();
                var usedSelection = new Selection(4, 1, 7, 10);
                var referenceSelection = new Selection(5, 1, 5, 10);
                IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null);

                var SUT = new ExtractMethodRuleUsedAfter();
                SUT.setValidFlag(ref flag, reference, usedSelection);

                Assert.AreEqual(0, flag);

            }

        }

    }

    [TestClass]
    public class ExtractMethodRuleUsedBeforeTests
    {
        [TestClass]
        public class WhenSetValidFlagIsCalledWhenTheReferenceIsUsedAfter : ExtractMethodRuleUsedBeforeTests
        {
            [TestMethod]
            [TestCategory("ExtractMethodRuleTests")]
            public void shouldNotSetFlag()
            {
                byte flag = new byte();
                var usedSelection = new Selection(4, 1, 7, 10);
                var referenceSelection = new Selection(8, 1, 8, 10);
                IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null);

                var SUT = new ExtractMethodRuleUsedBefore();
                SUT.setValidFlag(ref flag, reference, usedSelection);

                Assert.AreEqual(0, flag);

            }
        }
        [TestClass]
        public class WhenSetValidFlagIsCalledWhenTheReferenceIsUsedBefore : ExtractMethodRuleUsedBeforeTests
        {
            [TestMethod]
            [TestCategory("ExtractMethodRuleTests")]
            public void shouldSetFlagUsedBefore()
            {
                byte flag = new byte();
                var usedSelection = new Selection(4, 1, 7, 10);
                var referenceSelection = new Selection(3, 1, 3, 10);
                IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null);

                var SUT = new ExtractMethodRuleUsedBefore();
                SUT.setValidFlag(ref flag, reference, usedSelection);

                Assert.AreEqual((byte)ExtractMethodRuleFlags.UsedBefore, flag);

            }

        }
        [TestClass]
        public class WhenSetValidFlagIsCalledWhenTheReferenceIsInSelection : ExtractMethodRuleUsedAfterTests
        {
            [TestMethod]
            [TestCategory("ExtractMethodRuleTests")]
            public void shouldNotSetFlag()
            {
                byte flag = new byte();
                var usedSelection = new Selection(4, 1, 7, 10);
                var referenceSelection = new Selection(5, 1, 5, 10);
                IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null);

                var SUT = new ExtractMethodRuleUsedBefore();
                SUT.setValidFlag(ref flag, reference, usedSelection);

                Assert.AreEqual(0, flag);

            }

        }

    }

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
                    byte flag = new byte();
                    var usedSelection = new Selection(4, 1, 7, 10);
                    var referenceSelection = new Selection(6, 1, 6, 10);
                    var isAssigned = true;
                    IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null, isAssigned);

                    var SUT = new ExtractMethodRuleIsAssignedInSelection();
                    SUT.setValidFlag(ref flag, reference, usedSelection);

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
                    byte flag = new byte();
                    var usedSelection = new Selection(4, 1, 7, 10);
                    var referenceSelection = new Selection(6, 1, 6, 10);
                    var isAssigned = false;
                    IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null, isAssigned);

                    var SUT = new ExtractMethodRuleIsAssignedInSelection();
                    SUT.setValidFlag(ref flag, reference, usedSelection);

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
                    byte flag = new byte();
                    var usedSelection = new Selection(4, 1, 7, 10);
                    var referenceSelection = new Selection(3, 1, 3, 10);
                    var isAssigned = true;
                    IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null, isAssigned);

                    var SUT = new ExtractMethodRuleIsAssignedInSelection();
                    SUT.setValidFlag(ref flag, reference, usedSelection);

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
                    byte flag = new byte();
                    var usedSelection = new Selection(4, 1, 7, 10);
                    var referenceSelection = new Selection(9, 1, 9, 10);
                    var isAssigned = true;
                    IdentifierReference reference = new IdentifierReference(new QualifiedModuleName(), null, null, "a", referenceSelection, null, null, isAssigned);

                    var SUT = new ExtractMethodRuleIsAssignedInSelection();
                    SUT.setValidFlag(ref flag, reference, usedSelection);

                    Assert.AreEqual(0, flag);

                }
            }
        }
    }

    [TestClass]
    public class ExtractMethodModelTests
    {

        [TestClass]
        public class WhenExtractingFromASelection
        {
            #region hard coded data
            string inputCode = @"
Option explicit
Public Sub CodeWithDeclaration()
    Dim x as long
    Dim y as long
    Dim z as long

    x = 1 + 2
    DebugPrint x
    y = x + 1
    x = 2
    DebugPrint y

    z = x
    DebugPrint z

End Sub
Public Sub DebugPrint(byval g as long)
End Sub


";
            string selectedCode = @"
y = x + 1 
x = 2
Debug.Print y";

            string outputCode = @"
Public Sub NewVal( byval x as long)
    Dim y as long
    y = x + 1
    x = 2
    DebugPrint y
End Sub";
            #endregion

            List<IExtractMethodRule> emRules = new List<IExtractMethodRule>(){
                new ExtractMethodRuleInSelection(),
                new ExtractMethodRuleIsAssignedInSelection(),
                new ExtractMethodRuleUsedBefore(),
                new ExtractMethodRuleUsedAfter()};

            [TestMethod]
            [TestCategory("ExtractMethodModelTests")]
            public void shouldIncludeDeclarationsToRemoveInLinesToRemove()
            {

                QualifiedModuleName qualifiedModuleName;
                RubberduckParserState state;
                MockParser.ParseString(inputCode, out qualifiedModuleName, out state);
                var declarations = state.AllDeclarations;

                var selection = new Selection(10, 1, 12, 17);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var emr = new Mock<IExtractMethodRule>();
                var extractedMethod = new Mock<IExtractedMethod>();
                var extractedMethodModel = new ExtractMethodModel(emRules, extractedMethod.Object);
                extractedMethodModel.extract(declarations, qSelection.Value, selectedCode);

                var expected = new Selection(5, 9, 5, 10);

                Assert.IsTrue(extractedMethodModel.SelectionToRemove.Contains(expected), "The selectionToRemove should contain the Declaration being moved");

            }
            [TestMethod]
            [TestCategory("ExtractMethodModelTests")]
            public void shouldProvideTheSelectionOfLinesOfToRemove()
            {
                QualifiedModuleName qualifiedModuleName;
                RubberduckParserState state;
                MockParser.ParseString(inputCode, out qualifiedModuleName, out state);
                var declarations = state.AllDeclarations;

                var selection = new Selection(10, 1, 12, 17);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var emr = new Mock<IExtractMethodRule>();
                var extractedMethod = new Mock<IExtractedMethod>();
                var extractedMethodModel = new ExtractMethodModel(emRules,extractedMethod.Object);
                extractedMethodModel.extract(declarations, qSelection.Value, selectedCode);

                var selections = new List<Selection>() { new Selection(5, 9, 5, 10), selection };


                Assert.AreEqual(selections.Count(), extractedMethodModel.SelectionToRemove.Count(), "Selection To Remove doesn't have the right number of members");
                selections.ForEach(s => Assert.IsTrue(extractedMethodModel.SelectionToRemove.Contains(s), string.Format("selection {0} missing from actual SelectionToRemove", s)));

            }
            [TestMethod]
            [TestCategory("ExtractMethodModelTests")]
            public void shouldProvideTheExtractMethodCaller()
            {
                QualifiedModuleName qualifiedModuleName;
                RubberduckParserState state;
                MockParser.ParseString(inputCode, out qualifiedModuleName, out state);
                var declarations = state.AllDeclarations;

                var selection = new Selection(10, 1, 12, 17);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var emr = new Mock<IExtractMethodRule>();
                var extractedMethod = new Mock<IExtractedMethod>();
                var SUT = new ExtractMethodModel(emRules, extractedMethod.Object);


                var x = SUT.NewMethodCall;

                extractedMethod.Verify(em => em.NewMethodCall());


            }
            [TestMethod]
            [TestCategory("ExtractMethodModelTests")]
            public void shouldProvideTheNewExtractedMethod()
            {
                QualifiedModuleName qualifiedModuleName;
                RubberduckParserState state;
                MockParser.ParseString(inputCode, out qualifiedModuleName, out state);
                var declarations = state.AllDeclarations;

                var selection = new Selection(10, 1, 12, 17);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var emr = new Mock<IExtractMethodRule>();
                var extractedMethod = new Mock<IExtractedMethod>();
                var extractedMethodProc = new Mock<IExtractMethodProc>();
                var SUT = new ExtractMethodModel(emRules, extractedMethod.Object);

                var actual = SUT.NewExtractedMethod(extractedMethodProc.Object);
                extractedMethodProc.Verify( emp => emp.createProc(SUT));

            }
            [TestMethod]
            [TestCategory("ExtractMethodModelTests")]
            public void shouldProvideThePositionForTheMethodCall()
            {
                QualifiedModuleName qualifiedModuleName;
                RubberduckParserState state;
                MockParser.ParseString(inputCode, out qualifiedModuleName, out state);
                var declarations = state.AllDeclarations;

                var selection = new Selection(10, 1, 12, 17);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var emr = new Mock<IExtractMethodRule>();
                var extractedMethod = new Mock<IExtractedMethod>();
                var extractedMethodProc = new Mock<IExtractMethodProc>();
                var SUT = new ExtractMethodModel(emRules, extractedMethod.Object);
                SUT.extract(declarations, qSelection.Value, selectedCode);

                var expected = new Selection(9,1,9,1);
                var actual = SUT.PositionForMethodCall;

                Assert.AreEqual(expected, actual, "Call should have been at row " + expected + " but is at " + actual );


            }
            [TestMethod]
            [TestCategory("ExtractMethodModelTests")]
            public void shouldProvideThePositionForTheNewMethod()
            {
                QualifiedModuleName qualifiedModuleName;
                RubberduckParserState state;
                MockParser.ParseString(inputCode, out qualifiedModuleName, out state);
                var declarations = state.AllDeclarations;

                var selection = new Selection(10, 1, 12, 17);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var emr = new Mock<IExtractMethodRule>();
                var extractedMethod = new Mock<IExtractedMethod>();
                var extractedMethodProc = new Mock<IExtractMethodProc>();
                var SUT = new ExtractMethodModel(emRules, extractedMethod.Object);
                SUT.extract(declarations, qSelection.Value, selectedCode);

                var expected = new Selection(18,1,18,1);
                var actual = SUT.PositionForNewMethod;

                Assert.AreEqual(expected, actual, "Call should have been at row " + expected + " but is at " + actual );

            }

        }

        [TestClass]
        public class WhenExtracting
        {
            #region inputCode
            string inputCode = @"
Option explicit
Public Sub CodeWithDeclaration()
    Dim x as long
    Dim y as long
    Dim z as long

    x = 1 + 2
    DebugPrint x
    y = x + 1
    x = 2
    DebugPrint y

    z = x
    DebugPrint z

End Sub
Public Sub DebugPrint(byval g as long)
End Sub


";
            #endregion

            List<IExtractMethodRule> emRules = new List<IExtractMethodRule>(){
                new ExtractMethodRuleInSelection(),
                new ExtractMethodRuleIsAssignedInSelection(),
                new ExtractMethodRuleUsedBefore(),
                new ExtractMethodRuleUsedAfter()};

            [TestMethod]
            [TestCategory("ExtractMethodModelTests")]
            public void shouldCallEachExtractMethodRuleOnEachReference()
            {

                var selectedCode = @"
y = x + 1 
x = 2
Debug.Print y";


                QualifiedModuleName qualifiedModuleName;
                RubberduckParserState state;
                MockParser.ParseString(inputCode, out qualifiedModuleName, out state);
                var declarations = state.AllDeclarations;

                var selection = new Selection(10, 1, 12, 17);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var emr = new Mock<IExtractMethodRule>();
                var emRules = new List<IExtractMethodRule>() { emr.Object, emr.Object };
                var extractedMethod = new Mock<IExtractedMethod>();
                var extractedMethodModel = new ExtractMethodModel(emRules,extractedMethod.Object);
                extractedMethodModel.extract(declarations, qSelection.Value, selectedCode);
                var _byte = new Byte();

                //Verify each rule is called 9 times : 5 for x , 2 for y, 2 for z
                emr.Verify(r => r.setValidFlag(ref _byte, It.IsAny<IdentifierReference>(), It.IsAny<Selection>()), Times.Exactly(18));
            }

        }


        [TestClass]
        public class WhenLocalVariableConstantIsInternal
        {
            [TestMethod]
            [TestCategory("ExtractMethodModelTests")]
            public void shouldExcludeVariableInSignature()
            {

                #region inputCode
                var inputCode = @"
Option explicit
Public Sub CodeWithDeclaration()
    Dim x as long
    Dim y as long
    Dim z as long

    x = 1 + 2
    DebugPrint x
    y = x + 1
    x = 2
    DebugPrint y

    z = x
    DebugPrint z

End Sub
Public Sub DebugPrint(byval g as long)
End Sub


";

                var selectedCode = @"
y = x + 1 
x = 2
Debug.Print y";
                #endregion

                QualifiedModuleName qualifiedModuleName;
                RubberduckParserState state;
                MockParser.ParseString(inputCode, out qualifiedModuleName, out state);
                var declarations = state.AllDeclarations;

                var selection = new Selection(10, 1, 12, 17);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);
                var emRules = new List<IExtractMethodRule>() {
                    new ExtractMethodRuleInSelection(),
                    new ExtractMethodRuleIsAssignedInSelection(),
                    new ExtractMethodRuleUsedAfter(),
                    new ExtractMethodRuleUsedBefore()};
                var extractedMethod = new Mock<IExtractedMethod>();
                var extractedMethodModel = new ExtractMethodModel(emRules,extractedMethod.Object);
                extractedMethodModel.extract(declarations, qSelection.Value, selectedCode);

                var actual = extractedMethodModel.Method.NewMethodCall();
                var expected = "NewMethod x";

                Assert.AreEqual(expected, actual);
            }
        }
        [TestClass]
        public class WhenDeclarationsContainNoPreviousNewMethod
        {
            [TestMethod]
            [TestCategory("ExtractMethodModelTests")]
            public void shouldReturnNewMethod()
            {

                QualifiedModuleName qualifiedModuleName;
                RubberduckParserState state;
                var inputCode = @"
Option Explicit
Private Sub Foo()
    Dim x As Integer
    x = 1 + 2
End Sub";

                MockParser.ParseString(inputCode, out qualifiedModuleName, out state);
                var declarations = state.AllDeclarations;
                var selection = new Selection(5, 4, 5, 14);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var emRules = new List<IExtractMethodRule>() { };
                var extractedMethod = new Mock<IExtractedMethod>();
                var SUT = new ExtractMethodModel(emRules,extractedMethod.Object);
                SUT.extract(declarations, qSelection.Value, "x = 1 + 2");

                var actual = SUT.Method.MethodName;
                var expected = "NewMethod";

                Assert.AreEqual(expected, actual);

            }

        }

        [TestClass]
        public class WhenDeclarationsContainAPreviousNewMethod
        {
            [TestMethod]
            [TestCategory("ExtractMethodModelTests")]
            public void shouldReturnAnIncrementedMethodName()
            {

                QualifiedModuleName qualifiedModuleName;
                RubberduckParserState state;
                var inputCode = @"
Option Explicit
Private Sub Foo()
    Dim x As Integer
    x = 1 + 2
End Sub
Private Sub NewMethod
    dim a as string
    Debug.Print a
End Sub";

                MockParser.ParseString(inputCode, out qualifiedModuleName, out state);
                var declarations = state.AllDeclarations;
                var selection = new Selection(4, 4, 4, 14);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var emRules = new List<IExtractMethodRule>() { };
                var extractedMethod = new Mock<IExtractedMethod>();
                var SUT = new ExtractMethodModel(emRules,extractedMethod.Object);
                SUT.extract(declarations, qSelection.Value, "x = 1 + 2");

                var actual = SUT.Method.MethodName;
                var expected = "NewMethod1";

                Assert.AreEqual(expected, actual);

            }

        }

        [TestClass]
        public class WhenDeclarationsContainAPreviousUnOrderedNewMethod
        {
            [TestMethod]
            [TestCategory("ExtractMethodModelTests")]
            public void shouldReturnAnLeastNextMethod()
            {

                QualifiedModuleName qualifiedModuleName;
                RubberduckParserState state;
                var inputCode = @"
Option Explicit
Private Sub Foo()
    Dim x As Integer
    x = 1 + 2
End Sub
Private Sub NewMethod
    dim a as string
    Debug.Print a
End Sub
Private Sub NewMethod1
    dim a as string
    Debug.Print a
End Sub
Private Sub NewMethod4
    dim a as string
    Debug.Print a
End Sub";

                MockParser.ParseString(inputCode, out qualifiedModuleName, out state);
                var declarations = state.AllDeclarations;
                var selection = new Selection(4, 4, 4, 14);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var emRules = new List<IExtractMethodRule>() { };
                var extractedMethod = new Mock<IExtractedMethod>();
                var SUT = new ExtractMethodModel(emRules,extractedMethod.Object);
                SUT.extract(declarations, qSelection.Value, "x = 1 + 2");

                var actual = SUT.Method.MethodName;
                var expected = "NewMethod2";

                Assert.AreEqual(expected, actual);

            }

        }

    }
}
