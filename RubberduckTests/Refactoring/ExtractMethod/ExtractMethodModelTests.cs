using System;
using System.Diagnostics;
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
        public class when_declarations_contain_no_previous_newMethod
        {
            [TestMethod]
            public void should_return_newMethod()
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
                Debug.WriteLine(actual);
                var expected = "NewMethod";

                Assert.AreEqual(expected, actual);

            }

        }

        [TestClass]
        public class when_declarations_contain_a_previous_newMethod
        {
            [TestMethod]
            public void should_return_an_incremented_methodName()
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
        public class when_declarations_contain_a_previous_unordered_newMethod
        {
            [TestMethod]
            public void should_return_an_least_next_newMethod()
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
