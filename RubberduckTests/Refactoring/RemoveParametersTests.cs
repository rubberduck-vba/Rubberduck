using System;
using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System.Collections.Generic;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class RemoveParametersTests : InteractiveRefactoringTestBase<IRemoveParametersPresenter, RemoveParametersModel>
    {
        //TestCase arg1 => number of arguments in the Sub or Function call
        //TestCase arg2 => argument numbers to remove
        //Input and Expected results generated for each test
        [TestCase(1, "1")]
        [TestCase(2, "1")]
        [TestCase(2, "2")]
        [TestCase(2, "1,2")]
        [TestCase(3, "1")]
        [TestCase(3, "2")]
        [TestCase(3, "3")]
        [TestCase(3, "1,2")]
        [TestCase(3, "2,3")]   //Replicates Issue #4319
        [TestCase(3, "1,2,3")]
        [TestCase(6, "1,2")]
        [TestCase(6, "2,3")]
        [TestCase(6, "3,4")]
        [TestCase(6, "4,5")]
        [TestCase(6, "5,6")]
        [TestCase(6, "1,3,5")]
        [TestCase(6, "1,2,3")]
        [TestCase(6, "4,5,6")] //Replicates Issue #4319
        [TestCase(6, "1,5,6")] //Replicates Issue #4319
        [TestCase(6, "2,5,6")] //Replicates Issue #4319
        [TestCase(6, "3,5,6")] //Replicates Issue #4319
        [TestCase(6, "2,3,4,5,6")] //Replicates Issue #4319
        [TestCase(6, "1,2,3,4,5,6")]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_SignatureParamRemoval(int numParams, string paramsToRemove)
        {
            var preamble = "Private Sub Foo(";
            var input = preamble;
            for (var argNum = 1; argNum <= numParams; argNum++)
            {
                input = argNum == 1 
                    ? input + $"ar|g{argNum} As Long, " 
                    : input + $"arg{argNum} As Long, ";
            }
            input = input.Equals(preamble) 
                ? input 
                : input.Remove(input.Length - 2);
            input = input + ")";

            var paramsTR = paramsToRemove.Split(',');
            var userParamRemovalChoices = new List<int>();
            foreach (var idxString in paramsTR)
            {
                userParamRemovalChoices.Add(int.Parse(idxString) - 1);
            }

            var expect = preamble;
            for (var argNum = 1; argNum <= numParams; argNum++)
            {
                if (!userParamRemovalChoices.Contains(argNum - 1))
                {
                    expect = expect + $"arg{argNum} As Long, ";
                }
            }
            expect = expect.Equals(preamble) ? expect : expect.Remove(expect.Length - 2);
            expect = expect + ")";

            var inputCode =
$@"{input}
End Sub";

            var expectedCode =
$@"{expect}
End Sub";
            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        //TestCase arg1 => number of arguments in the Sub or Function call
        //TestCase arg2 => argument numbers to remove
        //Input and Expected results generated for each test.  This test generates references to modify as well
        [TestCase(2, "1")]
        [TestCase(3, "2,3")] //Replicates Issue #4319
        [TestCase(4, "3,4")] //Replicates Issue #4319
        [TestCase(5, "4,5")] //Replicates Issue #4319
        [TestCase(5, "3,4,5")] //Replicates Issue #4319
        [TestCase(5, "2,3,4,5")] //Replicates Issue #4319
        [TestCase(5, "1,2,3,4,5")]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_SignatureAndReferenceParamRemoval(int numParams, string paramsToRemove)
        {
            const string preamble = "Private Sub Foo(";
            const string refPreamble = "Foo ";
            var input = preamble;
            var refInput = refPreamble;
            for (var argNum = 1; argNum <= numParams; argNum++)
            {
                input = argNum == 1 
                    ? input + $"ar|g{argNum} As Long, " 
                    : input + $"arg{argNum} As Long, ";
                refInput = refInput + $"{argNum},";
            }
            input = input.Equals(preamble) 
                ? input 
                : input.Remove(input.Length - 2);
            input = input + ")";

            refInput = refInput.Remove(refInput.Length - 1);

            var paramsTR = paramsToRemove.Split(',');
            var userParamRemovalChoices = new List<int>();
            foreach (var idxString in paramsTR)
            {
                userParamRemovalChoices.Add(int.Parse(idxString) - 1);
            }

            var expect = preamble;
            var refExpect = refPreamble;
            for (var argNum = 1; argNum <= numParams; argNum++)
            {
                if (!userParamRemovalChoices.Contains(argNum - 1))
                {
                    expect = expect + $"arg{argNum} As Long, ";
                    refExpect = refExpect + $"{argNum},";
                }
            }
            expect = expect.Equals(preamble) ? expect : expect.Remove(expect.Length - 2);
            expect = expect + ")";
            refExpect = refExpect.Equals(refPreamble) ? refExpect : refExpect.Remove(refExpect.Length - 1);

            var inputCode =
$@"{input}
End Sub

Private Sub Bar()
    {refInput}
End Sub

Private Sub AnotherBar()
    {refInput}
End Sub";

            var expectedCode =
$@"{expect}
End Sub

Private Sub Bar()
    {refExpect}
End Sub

Private Sub AnotherBar()
    {refExpect}
End Sub";
            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveNamedParam_4params_4319()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal ar|g2 As String, ByVal arg3 As Double, ByVal arg4 As Double)
End Sub

Public Sub Goo()
    Foo arg2:=""test44"", arg3:=6.1, arg1:=3, arg4:= 8.2
End Sub
";
            const string expectedCode =
@"Public Sub Foo(ByVal arg1 As Integer)
End Sub

Public Sub Goo()
    Foo arg1:=3
End Sub
";

            var userParamRemovalChoices = new[] { 1, 2, 3 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveNamedParam_3params_4319()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal ar|g2 As String, ByVal arg3 As Double)
End Sub

Public Sub Goo()
    Foo arg1:=3, arg2:=""test44"", arg3:=6.1
End Sub
";
            const string expectedCode =
@"Public Sub Foo(ByVal arg1 As Integer)
End Sub

Public Sub Goo()
    Foo arg1:=3
End Sub
";

            var userParamRemovalChoices = new[] { 1, 2 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveNamedParam()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal ar|g2 As String, ByVal arg3 As Double)
End Sub

Public Sub Goo()
    Foo arg2:=""test44"", arg3:=6.1, arg1:=3
End Sub
";
            const string expectedCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Sub Goo()
    Foo arg2:=""test44"", arg1:=3
End Sub
";

            var userParamRemovalChoices = new[] { 2 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_CallerArgNameContainsOtherArgName()
        {
            const string inputCode =
@"Sub fo|o(a, b, c)
End Sub

Sub goo()
    foo asd, sdf, s
End Sub";

            const string expectedCode =
@"Sub foo(a, b)
End Sub

Sub goo()
    foo asd, sdf
End Sub";

            var userParamRemovalChoices = new[] { 2 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveLastFromFunction()
        {
            const string inputCode =
@"Private Function Foo(ByVal arg|1 As Integer, ByVal arg2 As String) As Boolean
End Function";

            const string expectedCode =
@"Private Function Foo(ByVal arg1 As Integer) As Boolean
End Function";

            var userParamRemovalChoices = new[] { 1 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveAllFromFunction()
        {
            const string inputCode =
@"Private Function Foo(ByVal arg|1 As Integer, ByVal arg2 As String) As Boolean
End Function";

            const string expectedCode =
@"Private Function Foo() As Boolean
End Function";

            var actual = RemoveParams(inputCode);
            Assert.AreEqual(expectedCode, actual);
        }

        [TestCase("Foo arg1, arg2", "Foo ")]
        [TestCase("test = Foo(arg1, arg2)", "test = Foo()")]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveAllFromFunction_UpdateCallReferences(string input, string expected)
        {
            var inputCode =
$@"Private Function Foo(ByVal ar|g1 As Integer, ByVal arg2 As String) As Boolean
End Function

Private Sub Goo(ByVal arg1 As Integer, ByVal arg2 As String)
    Dim test As Boolean
    {input}
End Sub
";

            var expectedCode =
$@"Private Function Foo() As Boolean
End Function

Private Sub Goo(ByVal arg1 As Integer, ByVal arg2 As String)
    Dim test As Boolean
    {expected}
End Sub
";

            var actual = RemoveParams(inputCode);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_UpdateCallReferences_4319()
        {
            const string inputCode =
@"Private Function Foo(ByVal ar|g1 As Integer, ByVal arg2 As String, ByVal arg3 As Long) As Boolean
End Function

Private Sub Goo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Long)
    Foo arg1, arg2, arg3
End Sub
";

            const string expectedCode =
@"Private Function Foo(ByVal arg1 As Integer) As Boolean
End Function

Private Sub Goo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Long)
    Foo arg1
End Sub
";

            var userParamRemovalChoices = new[] { 1,2 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_ParentIdentifierContainsParameterName()
        {
            const string inputCode =
@"Private Sub foo(a, |b, c, d, e, f, g)
End Sub

Private Sub goo()
    foo 1, 2, 3, 4, 5, 6, 7
End Sub";

            const string expectedCode =
@"Private Sub foo(a, b, e, g)
End Sub

Private Sub goo()
    foo 1, 2, 5, 7
End Sub";

            var userParamRemovalChoices = new[] {2,3,5};

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveFromGetter()
        {
            const string inputCode =
@"Private Property Get Foo(ByVal arg1 As Inte|ger) As Boolean
End Property";

            const string expectedCode =
@"Private Property Get Foo() As Boolean
End Property";

            var actual = RemoveParams(inputCode);
            Assert.AreEqual(expectedCode, actual);
        }

        //This scenario fails when run in Excel: 'Let' is not modified if 'Get' arg2 is removed
        //But, the MockParser returns a ParseError
        [Test,Ignore("MockParser unable to parse multiparam Let/Get (or Set/Get)")]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveFromMultiParamProperty()
        {
            const string inputCode =
@"Private Property Let Foo(ByVal arg1 As Integer, arg2 As Integer, arg3 As Integer, prop As Integer)
End Property;

Private Property Get Foo(ByVal arg1 As Integer, arg2 As Integer, arg3 As Integer) As Integer
End Property";

            const string expectedCode =
@"Private Property Let Foo(ByVal arg1 As Integer, arg3 As Integer, prop As Integer)
End Property;

Private Property Get Foo(ByVal arg1 As Integer, arg3 As Integer) As Integer
End Property";

            var userParamRemovalChoices = new[] { 1 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_QuickFix()
        {
            const string inputCode = @"
Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Property";

            const string expectedCode = @"
Private Property Set Foo(ByVal arg2 As String)
End Property";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode,  out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var parameter = state.AllUserDeclarations.SingleOrDefault(p =>
                    p.DeclarationType == DeclarationType.Parameter && p.IdentifierName == "arg1");
                if (parameter == null) { Assert.Inconclusive("Can't find 'arg1' parameter/target."); }

                var qualifiedSelection = parameter.QualifiedSelection;

                var refactoring = (RemoveParametersRefactoring)TestRefactoring(rewritingManager, state);
                refactoring.QuickFix(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveFirstParamFromSetter()
        {
            const string inputCode =
@"Private Property Set Foo(ByVal ar|g1 As Integer, ByVal arg2 As String)
End Property";

            const string expectedCode =
@"Private Property Set Foo(ByVal arg2 As String)
End Property";

            var userParamRemovalChoices = new[] { 0 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_ClientReferencesAreUpdated_FirstParam()
        {
            const string inputCode =
@"Private Sub Foo(ByVal ar|g1 As Integer, ByVal arg2 As String)
End Sub

Private Sub Bar()
    Foo 10, ""Hello""
End Sub
";

            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String)
End Sub

Private Sub Bar()
    Foo ""Hello""
End Sub
";

            var userParamRemovalChoices = new[] { 0 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_ClientReferencesAreUpdated_FirstParam_LineContinued()
        {
            const string inputCode =
@"Private Function Foo(ByVal arg1 As Integer, ByVal ar|g2 As String)
End Function

Private Sub Bar()
    Dim x As Variant    
    x = Foo _
        (10, ""Hello"")
End Sub
";

            const string expectedCode =
@"Private Function Foo(ByVal arg2 As String)
End Function

Private Sub Bar()
    Dim x As Variant    
    x = Foo _
        (""Hello"")
End Sub
";

            var userParamRemovalChoices = new[] { 0 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_ClientReferencesAreUpdated_FirstParam_ParensAroundCall()
        {
            const string inputCode =
@"Private Sub bar()
    Dim x As Integer
    Dim y As Integer
    y = foo(x, 42)
    Debug.Print y, x
End Sub

Private Function foo(ByRe|f a As Integer, ByVal b As Integer) As Integer
    a = b
    foo = a + b
End Function";
            const string expectedCode =
@"Private Sub bar()
    Dim x As Integer
    Dim y As Integer
    y = foo(42)
    Debug.Print y, x
End Sub

Private Function foo(ByVal b As Integer) As Integer
    a = b
    foo = a + b
End Function";

            var userParamRemovalChoices = new[] { 0 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_ClientReferencesAreUpdated_LastParam()
        {
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg|2 As String)
End Sub

Private Sub Bar()
    Foo 10, ""Hello""
End Sub
";

            const string expectedCode =
@"Private Sub Foo(ByVal arg1 As Integer)
End Sub

Private Sub Bar()
    Foo 10
End Sub
";

            var userParamRemovalChoices = new[] { 1 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_ClientReferencesAreUpdated_OtherModule()
        {
            //Input
            const string inputDeclaringCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub
";
            const string inputCallingCode = 
@"Private Sub Bar()
    Foo 10, ""Hello""
End Sub
";
            var selection = new Selection(1, 45, 1, 49);

            //Expectation
            const string expectedCallingCode =
@"Private Sub Bar()
    Foo 10
End Sub
";
            var paramIndices = new[] {1}.ToList();
            var presenterAction = StandardPresenterAction(paramIndices);
            var actualCode = RefactoredCode(
                "DeclarationModule",
                selection,
                presenterAction,
                null,
                false,
                ("DeclarationModule", inputDeclaringCode, ComponentType.StandardModule),
                ("CallingModule", inputCallingCode, ComponentType.StandardModule));

            Assert.AreEqual(expectedCallingCode, actualCode["CallingModule"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_ClientReferencesAreUpdated_ParamArray()
        {
            const string inputCode =
@"Sub Foo(ByVal arg1 As String, Param|Array arg2())
End Sub

Public Sub Goo(ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
    Foo ""test"", test1x, test2x, test3x, test4x, test5x, test6x
End Sub
";

            const string expectedCode =
            @"Sub Foo(ByVal arg1 As String)
End Sub

Public Sub Goo(ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
    Foo ""test""
End Sub
";

            var userParamRemovalChoices = new[] { 1 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_ParamArrayReference_Issue4319()
        {
            const string inputCode =
@"Sub Fo|o(ByVal arg1 As String, ByVal arg2 As String, ParamArray arg3())
End Sub

Public Sub Goo(ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
    Foo ""test"",""test2"", test1x, test2x, test3x, test4x, test5x, test6x
End Sub
";

            const string expectedCode =
@"Sub Foo(ByVal arg1 As String)
End Sub

Public Sub Goo(ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
    Foo ""test""
End Sub
";

            var userParamRemovalChoices = new[] { 1,2 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveLastParamFromSetter_NotAllowed()
        {
            //Input
            const string inputCode =
                @"Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27);

            RemoveParametersModel capturedModel = null;
            Func<RemoveParametersModel, RemoveParametersModel> presenterAction = model => 
            {
                capturedModel = model;
                return model;
            };

            RefactoredCode(inputCode, selection, presenterAction);

            Assert.AreEqual(1, capturedModel.Parameters.Count); // doesn't allow removing last param from setter
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveLastParamFromLetter_NotAllowed()
        {
            //Input
            const string inputCode =
                @"Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27);

            RemoveParametersModel capturedModel = null;
            Func<RemoveParametersModel, RemoveParametersModel> presenterAction = model =>
            {
                capturedModel = model;
                return model;
            };

            RefactoredCode(inputCode, selection, presenterAction);

            Assert.AreEqual(1, capturedModel.Parameters.Count); // doesn't allow removing last param from setter
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveFirstParamFromGetterAndSetter()
        {
            const string inputCode =
@"Private Property Get Foo(ByVal a|rg1 As Integer) As Object
End Property

Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As Object)
End Property

Private Sub Bar()
    Dim str As Object
    Set str = Foo(42)
    Set Foo(23) = str
End Sub";

            const string expectedCode =
@"Private Property Get Foo() As Object
End Property

Private Property Set Foo(ByVal arg2 As Object)
End Property

Private Sub Bar()
    Dim str As Object
    Set str = Foo()
    Set Foo() = str
End Sub";

            var userParamRemovalChoices = new[] { 0 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveFirstParamFromGetterAndLetter()
        {
            const string inputCode =
@"Private Property Get Foo(ByVal a|rg1 As Integer) As String
End Property

Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Property

Private Sub Bar()
    Dim str As String
    str = Foo(42)
    Foo(23) = str
End Sub";

            const string expectedCode =
@"Private Property Get Foo() As String
End Property

Private Property Let Foo(ByVal arg2 As String)
End Property

Private Sub Bar()
    Dim str As String
    str = Foo()
    Foo() = str
End Sub";

            var userParamRemovalChoices = new[] { 0 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_SignatureContainsOptionalParam()
        {
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, Optio|nal ByVal arg2 As String)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo arg1
End Sub";

            const string expectedCode =
@"Private Sub Foo(Optional ByVal arg2 As String)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo 
End Sub";

            var userParamRemovalChoices = new[] { 0 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveOptionalParam_LastParam()
        {
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, Optional ByVal arg2 As S|tring)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo arg1
    Foo 1, ""test""
End Sub";

            const string expectedCode =
@"Private Sub Foo(ByVal arg1 As Integer)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo arg1
    Foo 1
End Sub";

            var userParamRemovalChoices = new[] { 1 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveOptionalParam_FirstParam()
        {
            const string inputCode =
@"Private Sub F|oo(Optional ByVal arg1 As Integer, Optional ByVal arg2 As String)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo arg1
    Foo 1, ""test""
    Foo , ""test""
End Sub";

            const string expectedCode =
@"Private Sub Foo(Optional ByVal arg2 As String)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo 
    Foo ""test""
    Foo ""test""
End Sub";

            var userParamRemovalChoices = new[] { 0 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveOptionalParam_MiddleParam()
        {
            const string inputCode =
@"Private Sub Foo(Optional ByVal arg|1 As Integer, Optional ByVal arg2 As String, Optional ByVal arg3 As Integer)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo arg1
    Foo 1, ""test""
    Foo 1, ""test"", 3
    Foo 1, , 3
    Foo , ""test""
    Foo , ""test"", 3
    Foo ,, 3
End Sub";

            const string expectedCode =
@"Private Sub Foo(Optional ByVal arg1 As Integer, Optional ByVal arg3 As Integer)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo arg1
    Foo 1
    Foo 1, 3
    Foo 1, 3
    Foo 
    Foo , 3
    Foo ,3
End Sub";

            var userParamRemovalChoices = new[] { 1 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveOptionalParam_NamedArgument()
        {
            const string inputCode =
                @"Private Sub F|oo(Optional ByVal arg1 As Integer, Optional ByVal arg2 As String)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo arg1:=arg1
    Foo arg2:=""test""
End Sub";

            const string expectedCode =
                @"Private Sub Foo(Optional ByVal arg1 As Integer)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo arg1:=arg1
    Foo 
End Sub";

            var userParamRemovalChoices = new[] { 1 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveOptionalParam_RemovesTrailingMissingArguments()
        {
            const string inputCode =
                @"Private Sub F|oo(Optional ByVal arg1 As Integer = 0, Optional ByVal arg2 As Long = 23, Optional ByVal arg3 As String = vbNullString, Optional ByVal arg4 As Long = 42)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo ,,, 2
    Foo ,, arg4:=1
    Foo 4,,, 3
    Foo 3,, arg4:=1
End Sub";

            const string expectedCode =
                @"Private Sub Foo(Optional ByVal arg1 As Integer = 0, Optional ByVal arg2 As Long = 23, Optional ByVal arg3 As String = vbNullString)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo 
    Foo 
    Foo 4
    Foo 3
End Sub";

            var userParamRemovalChoices = new[] { 3 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveOptionalParam_RemovesTrailingMissingArgumentsForMultiplePatches()
        {
            const string inputCode =
                @"Private Sub F|oo(Optional ByVal arg1 As Integer = 0, Optional ByVal arg2 As Long = 23, Optional ByVal arg3 As String = vbNullString, Optional ByVal arg4 As Long = 42)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo , 4,, 2
End Sub";

            const string expectedCode =
                @"Private Sub Foo(Optional ByVal arg1 As Integer = 0, Optional ByVal arg3 As String = vbNullString)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo 
End Sub";

            var userParamRemovalChoices = new[] { 1, 3 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_SignatureOnMultipleLines()
        {
            const string inputCode =
@"Private Sub Foo(ByVal a|rg1 As Integer, _
    ByVal arg2 As String, _
    ByVal arg3 As Date)
End Sub";

            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String, _
    ByVal arg3 As Date)
End Sub";   // note: VBE removes excess spaces

            var userParamRemovalChoices = new[] { 0 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_SignatureOnMultipleLines_RemoveSecond()
        {
            const string inputCode =
@"Private Sub Foo(ByVal ar|g1 As Integer, _
    ByVal arg2 As String, _
    ByVal arg3 As Date)
End Sub";

            const string expectedCode =
@"Private Sub Foo(ByVal arg1 As Integer, _
    ByVal arg3 As Date)
End Sub";   // note: VBE removes excess spaces

            var userParamRemovalChoices = new[] { 1 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_SignatureOnMultipleLines_RemoveLast()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg|1 As Integer, _
    ByVal arg2 As String, _
    ByVal arg3 As Date)
End Sub";

            const string expectedCode =
@"Private Sub Foo(ByVal arg1 As Integer, _
    ByVal arg2 As String)
End Sub";   // note: VBE removes excess spaces

            var userParamRemovalChoices = new[] { 2 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_PassTargetIn()
        {
            const string inputCode =
@"Private Sub Foo(ByVal ar|g1 As Integer, _
    ByVal arg2 As String, _
    ByVal arg3 As Date)
End Sub";

            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String, _
    ByVal arg3 As Date)
End Sub";   // note: VBE removes excess spaces

            var userParamRemovalChoices = new[] { 0 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_CallOnMultipleLines()
        {
            const string inputCode =
@"Private Sub Foo(|ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date)
End Sub

Private Sub Goo(ByVal arg1 as Integer, ByVal arg2 As String, ByVal arg3 As Date)

    Foo arg1, _
        arg2, _
        arg3

End Sub
";

            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String, ByVal arg3 As Date)
End Sub

Private Sub Goo(ByVal arg1 as Integer, ByVal arg2 As String, ByVal arg3 As Date)

    Foo arg2, _
        arg3

End Sub
";   // note: IDE removes excess spaces

            var userParamRemovalChoices = new[] { 0 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_LastInterfaceParamRemoved()
        {
            //Input
            const string inputCode1 =
                @"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";

            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode1 =
                @"Public Sub DoSomething(ByVal a As Integer)
End Sub";
            const string expectedCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer)
End Sub";   // note: IDE removes excess spaces

            var paramIndices = new[] { 1 }.ToList();
            var presenterAction = StandardPresenterAction(paramIndices);
            var actualCode = RefactoredCode(
                "IClass1",
                selection,
                presenterAction,
                null,
                false,
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule));

            Assert.AreEqual(expectedCode1, actualCode["IClass1"]);
            Assert.AreEqual(expectedCode2, actualCode["Class1"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_LastInterfaceParamRemoved_ImplementationParamsHaveDifferentNames()
        {
            //Input
            const string inputCode1 =
                @"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal v1 As Integer, ByVal v2 As String)
End Sub";

            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode1 =
                @"Public Sub DoSomething(ByVal a As Integer)
End Sub";
            const string expectedCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal v1 As Integer)
End Sub";   // note: IDE removes excess spaces

            var paramIndices = new[] { 1 }.ToList();
            var presenterAction = StandardPresenterAction(paramIndices);
            var actualCode = RefactoredCode(
                "IClass1",
                selection,
                presenterAction,
                null,
                false,
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule));

            Assert.AreEqual(expectedCode1, actualCode["IClass1"]);
            Assert.AreEqual(expectedCode2, actualCode["Class1"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_LastInterfaceParamRemoved_ImplementationParamsHaveDifferentNames_TwoImplementations()
        {
            //Input
            const string inputCode1 =
                @"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal v1 As Integer, ByVal v2 As String)
End Sub";
            const string inputCode3 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal i As Integer, ByVal s As String)
End Sub";

            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode1 =
                @"Public Sub DoSomething(ByVal a As Integer)
End Sub";
            const string expectedCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal v1 As Integer)
End Sub";   // note: IDE removes excess spaces
            const string expectedCode3 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal i As Integer)
End Sub";   // note: IDE removes excess spaces

            var paramIndices = new[] { 1 }.ToList();
            var presenterAction = StandardPresenterAction(paramIndices);
            var actualCode = RefactoredCode(
                "IClass1",
                selection,
                presenterAction,
                null,
                false,
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode3, ComponentType.ClassModule));

            Assert.AreEqual(expectedCode1, actualCode["IClass1"]);
            Assert.AreEqual(expectedCode2, actualCode["Class1"]);
            Assert.AreEqual(expectedCode3, actualCode["Class2"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_InterfaceGetterParam_ImplementationLetAndSetParamRemoved()
        {
            //Input
            const string inputCode1 =
                @"Public Property Get Foo(ByVal a As Integer, ByVal b As String) As Variant
End Property

Public Property Let Foo(ByVal a As Integer, ByVal b As String, RHS As Variant)
End Property

Public Property Set Foo(ByVal a As Integer, ByVal b As String, RHS As Variant)
End Property";
            const string inputCode2 =
                @"Implements IClass1

Private Property Get IClass1_Foo(ByVal a As Integer, ByVal b As String) As Variant
End Property

Private Property Let IClass1_Foo(ByVal a As Integer, ByVal b As String, RHS As Variant)
End Property

Private Property Set IClass1_Foo(ByVal a As Integer, ByVal b As String, RHS As Variant)
End Property";

            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode1 =
                @"Public Property Get Foo(ByVal a As Integer) As Variant
End Property

Public Property Let Foo(ByVal a As Integer, RHS As Variant)
End Property

Public Property Set Foo(ByVal a As Integer, RHS As Variant)
End Property";
            const string expectedCode2 =
                @"Implements IClass1

Private Property Get IClass1_Foo(ByVal a As Integer) As Variant
End Property

Private Property Let IClass1_Foo(ByVal a As Integer, RHS As Variant)
End Property

Private Property Set IClass1_Foo(ByVal a As Integer, RHS As Variant)
End Property";   

            var paramIndices = new[] { 1 }.ToList();
            var presenterAction = StandardPresenterAction(paramIndices);
            var actualCode = RefactoredCode(
                "IClass1",
                selection,
                presenterAction,
                null,
                false,
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule));

            Assert.AreEqual(expectedCode1, actualCode["IClass1"]);
            Assert.AreEqual(expectedCode2, actualCode["Class1"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_LastEventParamRemoved()
        {
            //Input
            const string inputCode1 =
                @"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)";

            const string inputCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            var selection = new Selection(1, 15, 1, 15);

            //Expectation
            const string expectedCode1 =
                @"Public Event Foo(ByVal arg1 As Integer)";

            const string expectedCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal arg1 As Integer)
End Sub";   // note: IDE removes excess spaces

            var paramIndices = new[] { 1 }.ToList();
            var presenterAction = StandardPresenterAction(paramIndices);
            var actualCode = RefactoredCode(
                "Class1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule));
            
            Assert.AreEqual(expectedCode1, actualCode["Class1"]);
            Assert.AreEqual(expectedCode2, actualCode["Class2"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_LastEventParamRemoved_EventImplementationSelected()
        {
            //Input
            const string inputCode1 =
                @"Private WithEvents abc As Class2

Private Sub abc_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            const string inputCode2 =   
                @"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)";

            var selection = new Selection(3, 15, 3, 15);

            //Expectation
            const string expectedCode1 =
                @"Private WithEvents abc As Class2

Private Sub abc_Foo(ByVal arg1 As Integer)
End Sub";   // note: IDE removes excess spaces

            const string expectedCode2 =
                @"Public Event Foo(ByVal arg1 As Integer)";

            var paramIndices = new[] { 1 }.ToList();
            var presenterAction = StandardPresenterAction(paramIndices);
            var actualCode = RefactoredCode(
                "Class1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule));

            Assert.AreEqual(expectedCode1, actualCode["Class1"]);
            Assert.AreEqual(expectedCode2, actualCode["Class2"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_LastEventParamRemoved_ParamsHaveDifferentNames()
        {
            //Input
            const string inputCode1 =
                @"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)";

            const string inputCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal i As Integer, ByVal s As String)
End Sub";

            var selection = new Selection(1, 15, 1, 15);

            //Expectation
            const string expectedCode1 =
                @"Public Event Foo(ByVal arg1 As Integer)";

            const string expectedCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal i As Integer)
End Sub";   // note: IDE removes excess spaces

            var paramIndices = new[] { 1 }.ToList();
            var presenterAction = StandardPresenterAction(paramIndices);
            var actualCode = RefactoredCode(
                "Class1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule));

            Assert.AreEqual(expectedCode1, actualCode["Class1"]);
            Assert.AreEqual(expectedCode2, actualCode["Class2"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_LastEventParamRemoved_ParamsHaveDifferentNames_TwoHandlers()
        {
            //Input
            const string inputCode1 =
                @"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)";

            const string inputCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal i As Integer, ByVal s As String)
End Sub";
            const string inputCode3 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal v1 As Integer, ByVal v2 As String)
End Sub";

            var selection = new Selection(1, 15, 1, 15);

            //Expectation
            const string expectedCode1 =
                @"Public Event Foo(ByVal arg1 As Integer)";

            const string expectedCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal i As Integer)
End Sub";   // note: IDE removes excess spaces
            const string expectedCode3 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal v1 As Integer)
End Sub";   // note: IDE removes excess spaces

            var paramIndices = new[] { 1 }.ToList();
            var presenterAction = StandardPresenterAction(paramIndices);
            var actualCode = RefactoredCode(
                "Class1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
                ("Class3", inputCode3, ComponentType.ClassModule));

            Assert.AreEqual(expectedCode1, actualCode["Class1"]);
            Assert.AreEqual(expectedCode2, actualCode["Class2"]);
            Assert.AreEqual(expectedCode3, actualCode["Class3"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_LastInterfaceParamsRemoved_AcceptPrompt()
        {
            //Input
            const string inputCode1 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
                @"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";

            var selection = new Selection(3, 23, 3, 23);

            //Expectation
            const string expectedCode1 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer)
End Sub";   // note: IDE removes excess spaces

            const string expectedCode2 =
                @"Public Sub DoSomething(ByVal a As Integer)
End Sub";

            var paramIndices = new[] { 1 }.ToList();
            var presenterAction = StandardPresenterAction(paramIndices);
            var actualCode = RefactoredCode(
                "Class1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("IClass1", inputCode2, ComponentType.ClassModule));

            Assert.AreEqual(expectedCode1, actualCode["Class1"]);
            Assert.AreEqual(expectedCode2, actualCode["IClass1"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParams_RefactorDeclaration_FailsInvalidTarget()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var actualCode = RefactoredCode(inputCode, "TestModule1", DeclarationType.ProceduralModule, typeof(InvalidDeclarationTypeException));

            Assert.AreEqual(inputCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParams_RefactorDeclaration_FailsNoValidTargetSelected()
        {
            //Input
            const string inputCode =
                @"Private bar As Long

Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = Selection.Home;
            var actualCode = RefactoredCode(inputCode, selection, typeof(NoDeclarationForSelectionException));

            Assert.AreEqual(inputCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParams_PresenterIsNull()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {
                var qualifiedSelection = new QualifiedSelection(component.QualifiedModuleName, Selection.Home);
                var factory = new Mock<IRefactoringPresenterFactory>();
                factory.Setup(f => f.Create<IRemoveParametersPresenter, RemoveParametersModel>(It.IsAny<RemoveParametersModel>()))
                    .Returns(() => null); // resolves method overload resolution error

                var selectionService = MockedSelectionService();

                var refactoring = TestRefactoring(rewritingManager, state, factory.Object, selectionService);

                Assert.Throws<InvalidRefactoringPresenterException>(() => refactoring.Refactor(qualifiedSelection));

                Assert.AreEqual(inputCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParams_ModelIsNull()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            Func<RemoveParametersModel, RemoveParametersModel> presenterAction = model => null;

            var actualCode = RefactoredCode(inputCode, selection, presenterAction, typeof(InvalidRefactoringModelException));

            Assert.AreEqual(inputCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemovesParameterOfCorrectMethod()
        {
            //Input
            const string classCode =
                @"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            const string moduleCode =
                @"
Private Function Bar(ByVal i As Integer, ByVal s As String) As Class1
End Function

Private Sub Baz()
    Bar(42, ""Hello"").Foo 23, ""Hi""
End Sub";

            var selection = new Selection(2, 20, 2, 20);

            //Expectation
            const string expectedClassCode =
                @"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            const string expectedModuleCode =
                @"
Private Function Bar(ByVal i As Integer) As Class1
End Function

Private Sub Baz()
    Bar(42).Foo 23, ""Hi""
End Sub";

            var paramIndices = new[] { 1 }.ToList();
            var presenterAction = StandardPresenterAction(paramIndices);
            var actualCode = RefactoredCode(
                "Module1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", classCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.AreEqual(expectedClassCode, actualCode["Class1"]);
            Assert.AreEqual(expectedModuleCode, actualCode["Module1"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemovesParameterOfDefaultMemberAccess()
        {
            //Input
            const string classCode =
                @"
Public Function Foo(ByVal arg1 As Integer, ByVal arg2 As String) As String
Attribute Foo.VB_UserMemId = 0
End Function";
            var selection = new Selection(2, 51, 2, 51);

            const string moduleCode =
                @"
Private Sub Baz()
    Dim cls As Class1
    Set cls = new Class1
    Dim fooBar As Variant
    fooBar = cls(42, ""Hello"")
End Sub";
            
            //Expectation
            const string expectedClassCode =
                @"
Public Function Foo(ByVal arg1 As Integer) As String
Attribute Foo.VB_UserMemId = 0
End Function";

            const string expectedModuleCode =
                @"
Private Sub Baz()
    Dim cls As Class1
    Set cls = new Class1
    Dim fooBar As Variant
    fooBar = cls(42)
End Sub";

            var paramIndices = new[] { 1 }.ToList();
            var presenterAction = StandardPresenterAction(paramIndices);
            var actualCode = RefactoredCode(
                "Class1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", classCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.AreEqual(expectedClassCode, actualCode["Class1"]);
            Assert.AreEqual(expectedModuleCode, actualCode["Module1"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_ReplacesDictionaryAccessExpressions()
        {
            //Input
            const string classCode =
                @"
Public Function Foo(ByVal arg1 As String) As String
Attribute Foo.VB_UserMemId = 0
End Function";
            var selection = new Selection(2, 51, 2, 51);

            const string moduleCode =
                @"
Private Sub Baz()
    Dim cls As Class1
    Set cls = new Class1
    Dim fooBar As Variant
    fooBar = cls!Hello
End Sub";

            //Expectation
            const string expectedClassCode =
                @"
Public Function Foo() As String
Attribute Foo.VB_UserMemId = 0
End Function";

            const string expectedModuleCode =
                @"
Private Sub Baz()
    Dim cls As Class1
    Set cls = new Class1
    Dim fooBar As Variant
    fooBar = cls()
End Sub";

            var paramIndices = new[] { 0 }.ToList();
            var presenterAction = StandardPresenterAction(paramIndices);
            var actualCode = RefactoredCode(
                "Class1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", classCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.AreEqual(expectedClassCode, actualCode["Class1"]);
            Assert.AreEqual(expectedModuleCode, actualCode["Module1"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        [Ignore("This is postponed until after refactoring the refactoring setup and extracting a refactoring out of ExpandBangNotationQuickFix.")]
        public void RemoveParametersRefactoring_ExpandsWithDictionaryAccessExpressions()
        {
            //Input
            const string classCode =
                @"
Public Function Foo(ByVal arg1 As String) As String
Attribute Foo.VB_UserMemId = 0
End Function";
            var selection = new Selection(2, 51, 2, 51);

            const string moduleCode =
                @"
Private Sub Baz()
    Dim fooBar As Variant
    With New Class1
        fooBar = !Hello
    End With
End Sub";

            //Expectation
            const string expectedClassCode =
                @"
Public Function Foo() As String
Attribute Foo.VB_UserMemId = 0
End Function";

            const string expectedModuleCode =
                @"
Private Sub Baz()
    Dim fooBar As Variant
    With New Class1
        fooBar = .Foo
    End With
End Sub";

            var paramIndices = new[] { 0 }.ToList();
            var presenterAction = StandardPresenterAction(paramIndices);
            var actualCode = RefactoredCode(
                "Class1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", classCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.AreEqual(expectedClassCode, actualCode["Class1"]);
            Assert.AreEqual(expectedModuleCode, actualCode["Module1"]);
        }

        private string RemoveParams(string inputCode, Selection? selection = null, IEnumerable<int> paramIndices = null)
        {
            var (code, codeSelection) = CodeFromCodeStringLike(inputCode, selection);
            var presenterAction = StandardPresenterAction(paramIndices);

            return RefactoredCode(code, codeSelection, presenterAction);
        }

        private string RemoveParams(string inputCode, string targetName, DeclarationType declarationType, IEnumerable<int> paramIndices = null)
        {
            var (code, codeSelection) = CodeFromCodeStringLike(inputCode);
            var presenterAction = StandardPresenterAction(paramIndices);

            return RefactoredCode(code, targetName, declarationType, presenterAction);
        }

        private static Func<RemoveParametersModel, RemoveParametersModel> StandardPresenterAction(IEnumerable<int> paramIndices = null)
        {
            if (paramIndices == null)
            {
                return model =>
                {
                    model.RemoveParameters = model.Parameters;
                    return model;
                };
            }

            return model =>
            {
                var paramsToRemove = paramIndices
                    .Select(idx => model.Parameters[idx])
                    .ToList();

                model.RemoveParameters = paramsToRemove;
                return model;
            };
        }

        private static (string code, Selection selection) CodeFromCodeStringLike(string inputCode, Selection? selection = null)
        {
            var codeString = inputCode.ToCodeString();
            var derivedSelection = selection ?? codeString.CaretPosition.ToOneBased();

            return (codeString.Code, derivedSelection);
        }

        protected override IRefactoring TestRefactoring(
            IRewritingManager rewritingManager, 
            RubberduckParserState state,
            RefactoringUserInteraction<IRemoveParametersPresenter, RemoveParametersModel> userInteraction, 
            ISelectionService selectionService)
        {
            var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionService, state);
            var baseRefactoring = new RemoveParameterRefactoringAction(state, rewritingManager);
            return new RemoveParametersRefactoring(baseRefactoring, state, userInteraction, selectionService, selectedDeclarationProvider);
        }
    }
}
