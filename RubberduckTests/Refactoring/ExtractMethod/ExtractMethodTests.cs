using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.VBEditor;

namespace RubberduckTests.Refactoring.ExtractMethod
{

    [TestClass]
    public class ExtractedMethodTests
    {
        [TestClass]
        public class WhenAMethodIsDefined : ExtractedMethodTests
        {

            [TestCategory("ExtractedMethodTests")]
            [TestMethod]
            public void shouldReturnStringCorrectly()
            {
                var method = new ExtractedMethod();
                method.Accessibility = Accessibility.Private;
                method.MethodName = "Bar";
                method.ReturnValue = null;
                var insertCode = "Bar x";
                var newParam = new ExtractedParameter("Integer", ExtractedParameter.PassedBy.ByVal, "x");
                method.Parameters = new List<ExtractedParameter>() { newParam };

                var actual = method.NewMethodCall();
                Debug.Print(method.NewMethodCall());

                Assert.AreEqual(insertCode, actual);


            }
        }
        [TestClass]
        public class WhenDeclarationsContainNoPreviousNewMethod : ExtractedMethodTests
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

                var SUT = new ExtractedMethod();

                var expected = "NewMethod";
                //Act
                var actual = SUT.getNewMethodName(declarations);

                //Assert

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

                #region inputCode
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
                #endregion inputCode

                MockParser.ParseString(inputCode, out qualifiedModuleName, out state);
                var declarations = state.AllDeclarations;

                var SUT = new ExtractedMethod();

                var expected = "NewMethod1";
                //Act
                var actual = SUT.getNewMethodName(declarations);

                //Assert
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
                #region inputCode
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
                #endregion inputCode

                MockParser.ParseString(inputCode, out qualifiedModuleName, out state);
                var declarations = state.AllDeclarations;

                var SUT = new ExtractedMethod();

                var expected = "NewMethod2";
                //Act
                var actual = SUT.getNewMethodName(declarations);

                //Assert
                Assert.AreEqual(expected, actual);

            }

        }

    }

}
