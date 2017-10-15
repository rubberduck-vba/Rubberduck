using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.ExtractMethod
{
    [TestClass]
    public class ExtractMethodModelTests
    {

        #region variableInternalAndOnlyUsedInternally
        string internalVariable = @"
Option explicit
Public Sub CodeWithDeclaration()
    Dim x as long
    Dim z as long

    x = 1 + 2
    DebugPrint ""something""                   '8:
    Dim y as long  
    y = x + 1
    x = 2
    DebugPrint y                               '12:

    y = x
    DebugPrint y
    z = 1

End Sub
Public Sub DebugPrint(byval g as long)
End Sub
                                               '21:

";
        string selectedCode = @"
y = x + 1 
x = 2
Debug.Print y";

        string outputCode = @"
Public Sub NewVal( byval x as long, byval y as long)
    DebugPrint ""something""
    y = x + 1
    x = 2
    DebugPrint y
End Sub";
        #endregion

        List<IExtractMethodRule> emRules = new List<IExtractMethodRule>(){
                        new ExtractMethodRuleInSelection(),
                        new ExtractMethodRuleIsAssignedInSelection(),
                        new ExtractMethodRuleUsedBefore(),
                        new ExtractMethodRuleUsedAfter(),
                        new ExtractMethodRuleExternalReference()};

        [TestMethod]
        [TestCategory("ExtractMethodModelTests")]
        public void shouldClassifyDeclarations()
        {
            QualifiedModuleName qualifiedModuleName;
            var state = MockParser.ParseString(internalVariable, out qualifiedModuleName);
            var declarations = state.AllDeclarations;

            var selection = new Selection(8, 1, 12, 24);
            QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

            var extractedMethod = new Mock<IExtractedMethod>();
            var extractedMethodProc = new Mock<IExtractMethodProc>();
            var paramClassify = new Mock<IExtractMethodParameterClassification>();

            var SUT = new ExtractMethodModel(extractedMethod.Object, paramClassify.Object);
            SUT.extract(declarations, qSelection.Value, selectedCode);

            paramClassify.Verify(
                pc => pc.classifyDeclarations(qSelection.Value, It.IsAny<Declaration>()),
                Times.Exactly(3));
        }

        [TestClass]
        public class WhenExtractingFromASelection : ExtractedMethodTests
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
    y = x + 1       '10
    x = 2
    DebugPrint y    '12

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

            [TestClass]
            public class WhenTheSelectionIsNotWithinAMethod : WhenExtractingFromASelection
            {
                [TestMethod]
                [TestCategory("ExtractMethodModelTests")]
                [ExpectedException(typeof(InvalidOperationException))]
                public void shouldThrowAnException()
                {
                    QualifiedModuleName qualifiedModuleName;
                    var state = MockParser.ParseString(inputCode, out qualifiedModuleName);
                    var declarations = state.AllDeclarations;

                    var selection = new Selection(21, 1, 22, 17);
                    QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                    var emr = new Mock<IExtractMethodRule>();
                    var extractedMethod = new Mock<IExtractedMethod>();
                    var paramClassify = new Mock<IExtractMethodParameterClassification>();
                    var SUT = new ExtractMethodModel(extractedMethod.Object, paramClassify.Object);

                    //Act
                    SUT.extract(declarations, qSelection.Value, selectedCode);

                    //Assert
                    // ExpectedException
                }

            }

            [TestMethod]
            [TestCategory("ExtractMethodModelTests")]
            public void shouldProvideAListOfDimsNoLongerNeededInTheContainingMethod()
            {
                QualifiedModuleName qualifiedModuleName;
                var state = MockParser.ParseString(inputCode, out qualifiedModuleName);
                var declarations = state.AllDeclarations;

                var selection = new Selection(10, 1, 12, 17);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);
                var extractDecl = declarations.Where(x => x.IdentifierName.Equals("y"));

                var emr = new Mock<IExtractMethodRule>();
                var extractedMethod = new Mock<IExtractedMethod>();
                var paramClassify = new Mock<IExtractMethodParameterClassification>();
                paramClassify.Setup(pc => pc.DeclarationsToMove).Returns(extractDecl);
                var SUT = new ExtractMethodModel(extractedMethod.Object, paramClassify.Object);
                SUT.extract(declarations, qSelection.Value, selectedCode);

                Assert.AreEqual(1, SUT.DeclarationsToMove.Count());
                Assert.IsTrue(SUT.DeclarationsToMove.Contains(extractDecl.First()), "The selectionToRemove should contain the Declaration being moved");

            }

            [TestMethod]
            [TestCategory("ExtractMethodModelTests")]
            public void shouldProvideTheSelectionOfLinesOfToRemove()
            {
                QualifiedModuleName qualifiedModuleName;
                var state = MockParser.ParseString(inputCode, out qualifiedModuleName);
                var declarations = state.AllDeclarations;

                var selection = new Selection(10, 2, 12, 17);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var emr = new Mock<IExtractMethodRule>();
                var extractedMethod = new Mock<IExtractedMethod>();
                var paramClassify = new Mock<IExtractMethodParameterClassification>();
                var extractDecl = declarations.Where(x => x.IdentifierName.Equals("y"));
                paramClassify.Setup(pc => pc.DeclarationsToMove).Returns(extractDecl);
                var extractedMethodModel = new ExtractMethodModel(extractedMethod.Object, paramClassify.Object);

                //Act
                extractedMethodModel.extract(declarations, qSelection.Value, selectedCode);

                //Assert
                var actual = extractedMethodModel.RowsToRemove;
                var yDimSelection = new Selection(5, 9, 5, 10);
                var expected = new[] { selection, yDimSelection }
                    .Select(x => new Selection(x.StartLine, 1, x.EndLine, 1));
                var missing = expected.Except(actual);
                var extra = actual.Except(expected);
                missing.ToList().ForEach(x => Trace.WriteLine(string.Format("missing item {0}", x)));
                extra.ToList().ForEach(x => Trace.WriteLine(string.Format("extra item {0}", x)));

                Assert.AreEqual(expected.Count(), actual.Count(), "Selection To Remove doesn't have the right number of members");
                expected.ToList().ForEach(s => Assert.IsTrue(actual.Contains(s), string.Format("selection {0} missing from actual SelectionToRemove", s)));

            }

            [TestMethod]
            [TestCategory("ExtractMethodModelTests")]
            public void shouldProvideTheExtractMethodCaller()
            {
                QualifiedModuleName qualifiedModuleName;
                var state = MockParser.ParseString(inputCode, out qualifiedModuleName);
                var declarations = state.AllDeclarations;

                var selection = new Selection(10, 1, 12, 17);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var emr = new Mock<IExtractMethodRule>();
                var extractedMethod = new Mock<IExtractedMethod>();
                var paramClassify = new Mock<IExtractMethodParameterClassification>();
                var SUT = new ExtractMethodModel(extractedMethod.Object, paramClassify.Object);


                var x = SUT.NewMethodCall;

                extractedMethod.Verify(em => em.NewMethodCall());


            }

            [TestMethod]
            [TestCategory("ExtractMethodModelTests")]
            public void shouldProvideThePositionForTheMethodCall()
            {
                QualifiedModuleName qualifiedModuleName;
                var state = MockParser.ParseString(inputCode, out qualifiedModuleName);
                var declarations = state.AllDeclarations;

                var selection = new Selection(10, 1, 12, 17);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var emr = new Mock<IExtractMethodRule>();
                var extractedMethod = new Mock<IExtractedMethod>();
                var paramClassify = new Mock<IExtractMethodParameterClassification>();
                var SUT = new ExtractMethodModel(extractedMethod.Object, paramClassify.Object);
                SUT.extract(declarations, qSelection.Value, selectedCode);

                var expected = new Selection(10, 1, 10, 1);
                var actual = SUT.PositionForMethodCall;

                Assert.AreEqual(expected, actual, "Call should have been at row " + expected + " but is at " + actual);
            }

            [TestMethod]
            [TestCategory("ExtractMethodModelTests")]
            public void shouldProvideThePositionForTheNewMethod()
            {
                QualifiedModuleName qualifiedModuleName;
                var state = MockParser.ParseString(inputCode, out qualifiedModuleName);
                var declarations = state.AllDeclarations;

                var selection = new Selection(10, 1, 12, 17);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var emr = new Mock<IExtractMethodRule>();
                var extractedMethod = new Mock<IExtractedMethod>();
                var extractedMethodProc = new Mock<IExtractMethodProc>();
                var paramClassify = new Mock<IExtractMethodParameterClassification>();
                var SUT = new ExtractMethodModel(extractedMethod.Object, paramClassify.Object);
                //Act
                SUT.extract(declarations, qSelection.Value, selectedCode);

                //Assert
                var expected = new Selection(18, 1, 18, 1);
                var actual = SUT.PositionForNewMethod;

                Assert.AreEqual(expected, actual, "Call should have been at row " + expected + " but is at " + actual);

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
                var state = MockParser.ParseString(inputCode, out qualifiedModuleName);
                var declarations = state.AllDeclarations;

                var selection = new Selection(10, 1, 12, 17);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);
                var extractedMethod = new Mock<IExtractedMethod>();
                extractedMethod.Setup(em => em.NewMethodCall())
                    .Returns("NewMethod x");
                var paramClassify = new Mock<IExtractMethodParameterClassification>();

                var SUT = new ExtractMethodModel(extractedMethod.Object, paramClassify.Object);

                //Act
                SUT.extract(declarations, qSelection.Value, selectedCode);

                //Assert
                var actual = SUT.Method.NewMethodCall();
                var expected = "NewMethod x";

                Assert.AreEqual(expected, actual);
            }

        }

        [TestClass]
        public class WhenSplittingSelection
        {

            [TestMethod]
            [TestCategory("ExtractMethodModelTests")]
            public void shouldSplitTheCodeAroundTheDefinition()
            {

                #region inputCode
                var inputCode = @"
Option explicit
Public Sub CodeWithDeclaration()
    Dim x as long
    Dim y as long

    x = 1 + 2             
    DebugPrint x                      '8
    y = x + 1
    Dim z as long                     '10
    z = x
    DebugPrint z                      '12
    x = 2                             
    DebugPrint y


End Sub
Public Sub DebugPrint(byval g as long)
End Sub


";

                var selectedCode = @"
    DebugPrint x                      '8
    y = x + 1
    Dim z as long                     '10
    z = x
    DebugPrint z                      '12";
                #endregion

                #region whatItShouldLookLike
                /*
public sub NewMethod(ByVal x as long, ByRef y as long)
    Dim z as long
    DebugPrint x
    y = x + 1             
    z = x
    DebugPrint z
end sub
*/
                #endregion

                QualifiedModuleName qualifiedModuleName;
                var state = MockParser.ParseString(inputCode, out qualifiedModuleName);
                var declarations = state.AllDeclarations;
                var selection = new Selection(8, 1, 12, 50);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);
                var extractedMethod = new Mock<IExtractedMethod>();
                var paramClassify = new Mock<IExtractMethodParameterClassification>();
                var yDecl = declarations.Where(decl => decl.IdentifierName.Equals("z"));
                var SUT = new ExtractMethodModel(extractedMethod.Object, paramClassify.Object);
                //Act
                var actual = SUT.splitSelection(selection, declarations);
                //Assert
                var selection1 = new Selection(8, 1, 9, 1);
                var selection2 = new Selection(11, 1, 12, 1);

                Assert.AreEqual(selection1, actual.First(), "Top selection does not match");
                Assert.AreEqual(selection2, actual.Skip(1).First(), "Bottom selection does not match");

            }


        }

        [TestClass]
        public class GroupByConsecutiveTests
        {

            [TestMethod]
            public void testSelection()
            {
                IEnumerable<int> list = new List<int> { 2, 3, 4, 6, 8, 9, 10, 12, 13, 15 };
                var grouped = list.GroupByMissing(x => (x + 1), (x, y) => new Selection(x, 1, y, 1), (x, y) => y - x);
            }

            [TestMethod]
            [ExpectedException(typeof(ArgumentException))]
            public void isUnordered()
            {
                IEnumerable<int> list = new List<int> { 2, 3, 4, 6, 7, 9, 8, 12, 13, 15 };
                var grouped = list.GroupByMissing(x => (x + 1), (x, y) => Tuple.Create(x, y), (x, y) => y - x).ToList();
            }

            [TestMethod]
            public void emptyList()
            {
                IEnumerable<int> list = new List<int> { };
                var grouped = list.GroupByMissing(x => (x + 1), (x, y) => Tuple.Create(x, y), (x, y) => y - x).ToList();
                Assert.AreEqual(0, grouped.Count());
            }

            [TestMethod]
            public void listOfSingleItem()
            {
                IEnumerable<int> list = new List<int> { 2 };
                var grouped = list.GroupByMissing(x => (x + 1), (x, y) => Tuple.Create(x, y), (x, y) => y - x).ToList();

                foreach (var group in grouped)
                {
                    Trace.WriteLine(group);
                }
                Assert.IsTrue(grouped.Contains(Tuple.Create(2, 2)));
                Assert.AreEqual(1, grouped.Count());
            }

            [TestMethod]
            public void testingUsefulList()
            {
                IEnumerable<int> list = new List<int> { 2, 3, 4, 6, 8, 9, 10, 12, 13, 15 };
                var grouped = list.GroupByMissing(x => (x + 1), (x, y) => Tuple.Create(x, y), (x, y) => y - x).ToList();

                foreach (var group in grouped)
                {
                    Trace.WriteLine(group);
                }
                Assert.IsTrue(grouped.Contains(Tuple.Create(2, 4)));
                Assert.IsTrue(grouped.Contains(Tuple.Create(6, 6)));
                Assert.IsTrue(grouped.Contains(Tuple.Create(8, 10)));
                Assert.IsTrue(grouped.Contains(Tuple.Create(12, 13)));
                Assert.IsTrue(grouped.Contains(Tuple.Create(15, 15)));
                Assert.AreEqual(5, grouped.Count());
            }

        }

    }
}
