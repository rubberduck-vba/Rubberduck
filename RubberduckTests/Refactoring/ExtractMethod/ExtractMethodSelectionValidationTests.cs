using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.ExtractMethod
{
    [TestClass]
    public class ExtractMethodSelectionValidationTests
    {
        [TestClass]
        public class SpansSingleMethod : ExtractMethodSelectionValidationTests
        {
            [TestClass]
            public class WhenSelectionSpansMoreThanASingleMethod : SpansSingleMethod
            {

                [TestMethod]
                [TestCategory("ExtractMethodSelectionValidationTests")]
                public void shouldReturnFalse()
                {
                    var inputCode = @"
Option Explicit
Private Sub Foo()
    Dim x As Integer
    x = 1 + 2
End Sub


Private Function NewMethod
    dim a as string
    Debug.Print a
End Function


Private Sub NewMethod4
    dim a as string

    Debug.Print a
End Sub";

                    QualifiedModuleName qualifiedModuleName;
                    var state = MockParser.ParseString(inputCode, out qualifiedModuleName);
                    var declarations = state.AllDeclarations;
                    var selection = new Selection(4, 4, 10, 14);
                    QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                    var SUT = new ExtractMethodSelectionValidation(declarations);

                    var actual = SUT.withinSingleProcedure(qSelection.Value);
                    var expected = false;
                    Assert.AreEqual(expected, actual);

                }
            }
            [TestClass]
            public class WhenSeletionSpansWithinMethod : SpansSingleMethod
            {
                [TestMethod]
                [TestCategory("ExtractMethodSelectionValidationTests")]
                public void shouldReturnTrue()
                {
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


Private Sub NewMethod4
    dim a as string

    Debug.Print a
End Sub";

                    QualifiedModuleName qualifiedModuleName;
                    var state = MockParser.ParseString(inputCode, out qualifiedModuleName);
                    var declarations = state.AllDeclarations;
                    var selection = new Selection(4, 4, 5, 14);
                    QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                    var SUT = new ExtractMethodSelectionValidation(declarations);

                    var actual = SUT.withinSingleProcedure(qSelection.Value);

                    var expected = true;
                    Assert.AreEqual(expected, actual);

                }

                [TestMethod]
                [TestCategory("ExtractMethodSelectionValidationTests")]
                public void shouldReturnFalse()
                {
                    var inputCode = @"
Option Explicit
Private Sub Foo(byval a as long, _
                byval b as long)

    Dim x As Integer
    x = 1 + 2
End Sub


Private Sub NewMethod
End Sub";


                    QualifiedModuleName qualifiedModuleName;
                    var state = MockParser.ParseString(inputCode, out qualifiedModuleName);
                    var declarations = state.AllDeclarations;
                    var selection = new Selection(4, 4, 7, 14);
                    QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                    var SUT = new ExtractMethodSelectionValidation(declarations);

                    var actual = SUT.withinSingleProcedure(qSelection.Value);

                    var expected = false;
                    Assert.AreEqual(expected, actual);
                }
            }
        }
    }
}
