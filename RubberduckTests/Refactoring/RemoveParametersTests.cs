using System;
using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.UI.Refactorings.RemoveParameters;
using System.Collections.Generic;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class RemoveParametersTests
    {
        //TestCase arg1 => number of arguments in the Sub or Function call
        //TestCase arg2 => arguement numbers to remove
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
                input = argNum == 1 ? input + $"ar|g{argNum} As Long, " : input + $"arg{argNum} As Long, ";
            }
            input = input.Equals(preamble) ? input : input.Remove(input.Length - 2);
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

            string inputCode =
$@"{input}
End Sub";

            string expectedCode =
$@"{expect}
End Sub";
            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        //TestCase arg1 => number of arguments in the Sub or Function call
        //TestCase arg2 => arguement numbers to remove
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
            var preamble = "Private Sub Foo(";
            var refPreamble = "Foo ";
            var input = preamble;
            var refInput = refPreamble;
            for (var argNum = 1; argNum <= numParams; argNum++)
            {
                input = argNum == 1 ? input + $"ar|g{argNum} As Long, " : input + $"arg{argNum} As Long, ";
                refInput = refInput + $"{argNum},";
            }
            input = input.Equals(preamble) ? input : input.Remove(input.Length - 2);
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

            string inputCode =
$@"{input}
End Sub

Private Sub Bar()
    {refInput}
End Sub

Private Sub AnotherBar()
    {refInput}
End Sub";

            string expectedCode =
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

            var userParamRemovalChoices = new int[] { 1, 2, 3 };

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

            var userParamRemovalChoices = new int[] { 1, 2 };

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

            var userParamRemovalChoices = new int[] { 2 };

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

            var userParamRemovalChoices = new int[] { 2 };

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

            var userParamRemovalChoices = new int[] { 1 };

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
            string inputCode =
$@"Private Function Foo(ByVal ar|g1 As Integer, ByVal arg2 As String) As Boolean
End Function

Private Sub Goo(ByVal arg1 As Integer, ByVal arg2 As String)
    Dim test As Boolean
    {input}
End Sub
";

            string expectedCode =
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

            var userParamRemovalChoices = new int[] { 1,2 };

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

            var userParamRemovalChoices = new int[] {2,3,5};

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

            var userParamRemovalChoices = new int[] { 1 };

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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode,  out IVBComponent component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var parameter = state.AllUserDeclarations.SingleOrDefault(p =>
                    p.DeclarationType == DeclarationType.Parameter && p.IdentifierName == "arg1");
                if (parameter == null) { Assert.Inconclusive("Can't find 'arg1' parameter/target."); }

                var qualifiedSelection = parameter.QualifiedSelection;

                //Specify Param(s) to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.QuickFix(state, qualifiedSelection);

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

            var userParamRemovalChoices = new int[] { 0 };

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

            var userParamRemovalChoices = new int[] { 0 };

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

            var userParamRemovalChoices = new int[] { 0 };

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

            var userParamRemovalChoices = new int[] { 0 };

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

            var userParamRemovalChoices = new int[] { 1 };

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
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCallingCode =
@"Private Sub Bar()
    Foo 10
End Sub
";

            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("DeclarationModule", ComponentType.StandardModule, inputDeclaringCode, selection)
                .AddComponent("CallingModule", ComponentType.StandardModule, inputCallingCode);
            projectBuilder.AddProjectToVbeBuilder();
            var vbe = vbeBuilder.Build();
            var declaringComponent = projectBuilder.MockComponents[0].Object;
            var callingComponent = projectBuilder.MockComponents[1].Object; 

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(declaringComponent), selection);

                //Specify Param(s) to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.RemoveParameters = new[] { model.Parameters[1] }.ToList();

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);
                var resultCallingCode = callingComponent.CodeModule.Content();

                Assert.AreEqual(expectedCallingCode, resultCallingCode);
            }
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

            var userParamRemovalChoices = new int[] { 1 };

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

            var userParamRemovalChoices = new int[] { 1,2 };

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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out IVBComponent component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var model = new RemoveParametersModel(state, qualifiedSelection, null);

                Assert.AreEqual(1, model.Parameters.Count); // doesn't allow removing last param from setter
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out IVBComponent component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var model = new RemoveParametersModel(state, qualifiedSelection, null);

                Assert.AreEqual(1, model.Parameters.Count); // doesn't allow removing last param from letter
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveFirstParamFromGetterAndSetter()
        {
            const string inputCode =
@"Private Property Get Foo(ByVal a|rg1 As Integer) As String
End Property

Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Property";

            const string expectedCode =
@"Private Property Get Foo() As String
End Property

Private Property Set Foo(ByVal arg2 As String)
End Property";

            var userParamRemovalChoices = new int[] { 0 };

            var actual = RemoveParams(inputCode, paramIndices: userParamRemovalChoices);
            Assert.AreEqual(expectedCode, actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveFirstParamFromGetterAndLetter()
        {
            const string inputCode =
@"Private Property Get Foo(ByVal a|rg1 As Integer)
End Property

Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Property";

            const string expectedCode =
@"Private Property Get Foo()
End Property

Private Property Let Foo(ByVal arg2 As String)
End Property";

            var userParamRemovalChoices = new int[] { 0 };

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

            var userParamRemovalChoices = new int[] { 0 };

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

            var userParamRemovalChoices = new int[] { 1 };

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

            var userParamRemovalChoices = new int[] { 0 };

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

            var userParamRemovalChoices = new int[] { 1 };

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

            var userParamRemovalChoices = new int[] { 0 };

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

            var userParamRemovalChoices = new int[] { 1 };

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

            var userParamRemovalChoices = new int[] { 2 };

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

            var userParamRemovalChoices = new int[] { 0 };

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

            var userParamRemovalChoices = new int[] { 0 };

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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);

                var module1 = project.Object.VBComponents[0].CodeModule;
                var module2 = project.Object.VBComponents[1].CodeModule;

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.RemoveParameters = new[] { model.Parameters[1] }.ToList();

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode1, module1.Content());
                Assert.AreEqual(expectedCode2, module2.Content());
            }
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);

                var module1 = project.Object.VBComponents[0].CodeModule;
                var module2 = project.Object.VBComponents[1].CodeModule;

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.RemoveParameters = new[] { model.Parameters[1] }.ToList();

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode1, module1.Content());
                Assert.AreEqual(expectedCode2, module2.Content());
            }
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode3)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);

                var module1 = project.Object.VBComponents[0].CodeModule;
                var module2 = project.Object.VBComponents[1].CodeModule;
                var module3 = project.Object.VBComponents[2].CodeModule;

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.RemoveParameters = new[] { model.Parameters[1] }.ToList();

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode1, module1.Content());
                Assert.AreEqual(expectedCode2, module2.Content());
                Assert.AreEqual(expectedCode3, module3.Content());
            }
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);

                var module1 = project.Object.VBComponents[0].CodeModule;
                var module2 = project.Object.VBComponents[1].CodeModule;

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.RemoveParameters = new[] { model.Parameters[1] }.ToList();

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode1, module1.Content());
                Assert.AreEqual(expectedCode2, module2.Content());
            }
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);

                var module1 = project.Object.VBComponents[0].CodeModule;
                var module2 = project.Object.VBComponents[1].CodeModule;

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.RemoveParameters = new[] { model.Parameters.Last() }.ToList();

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode1, module1.Content());
                Assert.AreEqual(expectedCode2, module2.Content());
            }
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);

                var module1 = project.Object.VBComponents[0].CodeModule;
                var module2 = project.Object.VBComponents[1].CodeModule;

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.RemoveParameters = new[] { model.Parameters[1] }.ToList();

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode1, module1.Content());
                Assert.AreEqual(expectedCode2, module2.Content());
            }
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class3", ComponentType.ClassModule, inputCode3)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);

                var module1 = project.Object.VBComponents[0].CodeModule;
                var module2 = project.Object.VBComponents[1].CodeModule;
                var module3 = project.Object.VBComponents[2].CodeModule;

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.RemoveParameters = new[] { model.Parameters[1] }.ToList();

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode1, module1.Content());
                Assert.AreEqual(expectedCode2, module2.Content());
                Assert.AreEqual(expectedCode3, module3.Content());
            }
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);

                var module1 = project.Object.VBComponents[0].CodeModule;
                var module2 = project.Object.VBComponents[1].CodeModule;

                var messageBox = new Mock<IMessageBox>();
                messageBox.Setup(m => m.ConfirmYesNo(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>())).Returns(true);

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, messageBox.Object);
                model.RemoveParameters = new[] { model.Parameters[1] }.ToList();

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode1, module1.Content());
                Assert.AreEqual(expectedCode2, module2.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_LastInterfaceParamRemoved_RejectPrompt()
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);

                var messageBox = new Mock<IMessageBox>();
                messageBox.Setup(m => m.ConfirmYesNo(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>())).Returns(false);

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, messageBox.Object);
                Assert.IsNull(model.TargetDeclaration);
            }
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
            var selection = new Selection(1, 23, 1, 27);

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out IVBComponent component, selection);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //set up model
                var model = new RemoveParametersModel(state, qualifiedSelection, null);

                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);

                try
                {
                    refactoring.Refactor(
                        model.Declarations.FirstOrDefault(
                            i => i.DeclarationType == DeclarationType.ProceduralModule));
                }
                catch (ArgumentException e)
                {
                    Assert.AreEqual("Invalid declaration type", e.Message);
                    return;
                }

                Assert.Fail();
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out IVBComponent component);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var factory = new RemoveParametersPresenterFactory(vbe.Object, null, state, null);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory);
                refactoring.Refactor();

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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out IVBComponent component, selection);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Param(s) to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(inputCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void Presenter_Accept_AutoMarksSingleParamAsRemoved()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer)
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out IVBComponent component, selection);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var model = new RemoveParametersModel(state, qualifiedSelection, new Mock<IMessageBox>().Object);
                model.RemoveParameters = new[] { model.Parameters[0] }.ToList();

                var factory = new RemoveParametersPresenterFactory(vbe.Object, null, state, null);

                var presenter = factory.Create();

                Assert.IsTrue(model.Parameters[0].Declaration.Equals(presenter.Show().Parameters[0].Declaration));
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void Presenter_ParameterlessTargetReturnsNullModel()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            var builder = new MockVbeBuilder();
            var projectBuilder = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("Module1", ComponentType.StandardModule, inputCode, selection);
            var project = projectBuilder.Build();
            builder.AddProject(project);
            var vbe = builder.Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var factory = new RemoveParametersPresenterFactory(vbe.Object, null, state, new Mock<IMessageBox>().Object);
                var presenter = factory.Create();

                Assert.AreEqual(null, presenter.Show());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void Factory_NullSelectionNullReturnsNullPresenter()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
End Sub";

            var builder = new MockVbeBuilder();
            var projectBuilder = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("Module1", ComponentType.StandardModule, inputCode);
            var project = projectBuilder.Build();
            builder.AddProject(project);
            var vbe = builder.Build();

            vbe.Setup(v => v.ActiveCodePane).Returns((ICodePane)null);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var factory = new RemoveParametersPresenterFactory(vbe.Object, null, state, null);

                Assert.IsNull(factory.Create());
            }
        }

        private string RemoveParams(string inputCode, bool passInTarget = false, Selection? selection = null, IEnumerable<int> paramIndices = null)
        {
            var codeString = inputCode.ToCodeString();
            if (!selection.HasValue)
            {
                Selection? derivedSelect = codeString.CaretPosition.ToOneBased();

                if (!derivedSelect.HasValue)
                {
                    Assert.Fail($"Unable to derive user selection for test");
                }
                selection = derivedSelect;
            }

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(codeString.Code, out IVBComponent component, selection.Value);
            var result = string.Empty;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection.Value);

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                if (paramIndices is null)
                {
                    model.RemoveParameters = model.Parameters;
                }
                else
                {
                    var paramsToRemove = new List<Parameter>();
                    foreach (var idx in paramIndices)
                    {
                        paramsToRemove.Add(model.Parameters[idx]);
                    }
                    model.RemoveParameters = paramsToRemove;
                }

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                if (passInTarget)
                {
                    refactoring.Refactor(model.TargetDeclaration);
                }
                else
                {
                    refactoring.Refactor(qualifiedSelection);
                }
                result = component.CodeModule.Content();
            }
            return result;
        }

        #region setup
        private static Mock<IRefactoringPresenterFactory<IRemoveParametersPresenter>> SetupFactory(RemoveParametersModel model)
        {
            var presenter = new Mock<IRemoveParametersPresenter>();
            presenter.Setup(p => p.Show()).Returns(model);

            var factory = new Mock<IRefactoringPresenterFactory<IRemoveParametersPresenter>>();
            factory.Setup(f => f.Create()).Returns(presenter.Object);
            return factory;
        }

        #endregion
    }
}
