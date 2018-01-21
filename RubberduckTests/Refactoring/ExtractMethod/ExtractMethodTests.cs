using System.Collections.Generic;
using System.Diagnostics;
using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.ExtractMethod
{

    [TestFixture]
    public class ExtractedMethodTests
    {
        [TestFixture]
        public class WhenAMethodIsDefined : ExtractedMethodTests
        {

            [Category("ExtractedMethodTests")]
            [Test]
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
        [TestFixture]
        public class WhenDeclarationsContainNoPreviousNewMethod : ExtractedMethodTests
        {
            [Test]
            [Category("ExtractMethodModelTests")]
            public void shouldReturnNewMethod()
            {
                var inputCode = @"
Option Explicit
Private Sub Foo()
    Dim x As Integer
    x = 1 + 2
End Sub";

                QualifiedModuleName qualifiedModuleName;
                using (var state = MockParser.ParseString(inputCode, out qualifiedModuleName))
                {
                    var declarations = state.AllDeclarations;

                    var SUT = new ExtractedMethod();

                    var expected = "NewMethod";
                    //Act
                    var actual = SUT.getNewMethodName(declarations);

                    //Assert

                    Assert.AreEqual(expected, actual);

                }
            }

        }

        [TestFixture]
        public class WhenDeclarationsContainAPreviousNewMethod
        {
            [Test]
            [Category("ExtractMethodModelTests")]
            public void shouldReturnAnIncrementedMethodName()
            {
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

                QualifiedModuleName qualifiedModuleName;
                using (var state = MockParser.ParseString(inputCode, out qualifiedModuleName))
                {
                    var declarations = state.AllDeclarations;

                    var SUT = new ExtractedMethod();

                    var expected = "NewMethod1";
                    //Act
                    var actual = SUT.getNewMethodName(declarations);

                    //Assert
                    Assert.AreEqual(expected, actual);

                }
            }

        }

        [TestFixture]
        public class WhenDeclarationsContainAPreviousUnOrderedNewMethod
        {
            [Test]
            [Category("ExtractMethodModelTests")]
            public void shouldReturnAnLeastNextMethod()
            {
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

                QualifiedModuleName qualifiedModuleName;
                using (var state = MockParser.ParseString(inputCode, out qualifiedModuleName))
                {
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

}
