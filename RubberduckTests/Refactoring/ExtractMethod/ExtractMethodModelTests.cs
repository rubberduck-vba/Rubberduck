using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.VBEditor;

namespace RubberduckTests.Refactoring.ExtractMethod
{
    [TestClass]
    public class ExtractMethodModelTests
    {

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

    x = 1 + 2
    Debug.Print x
    y = x + 1
    Debug.Print y

End Sub
";

                var selectedCode = @"
y = x + 1 
Debug.Print y";
                #endregion

                QualifiedModuleName qualifiedModuleName;
                RubberduckParserState state;
                MockParser.ParseString(inputCode, out qualifiedModuleName, out state);
                var declarations = state.AllDeclarations;

                var selection = new Selection(9, 1, 10, 17);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);
                var extractedMethodModel = new ExtractMethodModel(declarations, qSelection.Value, selectedCode);

                var actual = extractedMethodModel.Method;
                var expected = "NewMethod  x";
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
                var selection = new Selection(4, 4, 4, 14);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var SUT = new ExtractMethodModel(declarations, qSelection.Value, "x = 1 + 2");

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

                var SUT = new ExtractMethodModel(declarations, qSelection.Value, "x = 1 + 2");

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

                var SUT = new ExtractMethodModel(declarations, qSelection.Value, "x = 1 + 2");

                var actual = SUT.Method.MethodName;
                var expected = "NewMethod2";

                Assert.AreEqual(expected, actual);

            }

        }
    }
}
