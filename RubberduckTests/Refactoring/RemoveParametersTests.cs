using System;
using System.Linq;
using System.Windows.Forms;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class RemoveParametersTests
    {
        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveBothParams()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters.ForEach(arg => arg.IsRemoved = true);

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveOnlyParam()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters.ForEach(arg => arg.IsRemoved = true);

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveFirstParam()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal arg2 As String)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Param(s) to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[0].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveSecondParam()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Param(s) to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[1].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveNamedParam()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Double)
End Sub

Public Sub Goo()
    Foo arg2:=""test44"", arg3:=6.1, arg1:=3
End Sub
";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Sub Goo()
    Foo arg2:=""test44"", arg1:=3
End Sub
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Param(s) to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[2].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_CallerArgNameContainsOtherArgName()
        {
            //Input
            const string inputCode =
                @"Sub foo(a, b, c)

End Sub

Sub goo()
    foo asd, sdf, s
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Sub foo(a, b)

End Sub

Sub goo()
    foo asd, sdf
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Param(s) to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[2].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveLastFromFunction()
        {
            //Input
            const string inputCode =
                @"Private Function Foo(ByVal arg1 As Integer, ByVal arg2 As String) As Boolean
End Function";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Function Foo(ByVal arg1 As Integer) As Boolean
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Param(s) to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[1].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveAllFromFunction()
        {
            //Input
            const string inputCode =
                @"Private Function Foo(ByVal arg1 As Integer, ByVal arg2 As String) As Boolean
End Function";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Function Foo() As Boolean
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Param(s) to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters.ForEach(p => p.IsRemoved = true);

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveAllFromFunction_UpdateCallReferences()
        {
            //Input
            const string inputCode =
                @"Private Function Foo(ByVal arg1 As Integer, ByVal arg2 As String) As Boolean
End Function

Private Sub Goo(ByVal arg1 As Integer, ByVal arg2 As String)
    Foo arg1, arg2
End Sub
";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Function Foo() As Boolean
End Function

Private Sub Goo(ByVal arg1 As Integer, ByVal arg2 As String)
    Foo 
End Sub
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Param(s) to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters.ForEach(p => p.IsRemoved = true);

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_ParentIdentifierContainsParameterName()
        {
            //Input
            const string inputCode =
                @"Private Sub foo(a, b, c, d, e, f, g)
End Sub

Private Sub goo()
    foo 1, 2, 3, 4, 5, 6, 7
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Sub foo(a, b, e, g)
End Sub

Private Sub goo()
    foo 1, 2, 5, 7
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters.ElementAt(2).IsRemoved = true;
                model.Parameters.ElementAt(3).IsRemoved = true;
                model.Parameters.ElementAt(5).IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveFromGetter()
        {
            //Input
            const string inputCode =
                @"Private Property Get Foo(ByVal arg1 As Integer) As Boolean
End Property";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Property Get Foo() As Boolean
End Property";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Param(s) to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters.ForEach(p => p.IsRemoved = true);

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_QuickFix()
        {
            //Input
            const string inputCode = @"
Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Property";

            //Expectation
            const string expectedCode = @"
Private Property Set Foo(ByVal arg2 As String)
End Property";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
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
            //Input
            const string inputCode =
                @"Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Property";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Property Set Foo(ByVal arg2 As String)
End Property";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Param(s) to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[0].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_ClientReferencesAreUpdated_FirstParam()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Private Sub Bar()
    Foo 10, ""Hello""
End Sub
";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal arg2 As String)
End Sub

Private Sub Bar()
    Foo ""Hello""
End Sub
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Param(s) to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[0].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_ClientReferencesAreUpdated_FirstParam_ParensAroundCall()
        {
            //Input
            const string inputCode =
                @"Private Sub bar()
    Dim x As Integer
    Dim y As Integer
    y = foo(x, 42)
    Debug.Print y, x
End Sub

Private Function foo(ByRef a As Integer, ByVal b As Integer) As Integer
    a = b
    foo = a + b
End Function";
            var selection = new Selection(8, 20, 8, 20);

            //Expectation
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Param(s) to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[0].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_ClientReferencesAreUpdated_LastParam()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Private Sub Bar()
    Foo 10, ""Hello""
End Sub
";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal arg1 As Integer)
End Sub

Private Sub Bar()
    Foo 10
End Sub
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Param(s) to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[1].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_ClientReferencesAreUpdated_ParamArray()
        {
            //Input
            const string inputCode =
                @"Sub Foo(ByVal arg1 As String, ParamArray arg2())
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
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Param(s) to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[1].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
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
            //Input
            const string inputCode =
                @"Private Property Get Foo(ByVal arg1 As Integer)
End Property

Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Property";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Property Get Foo()
End Property

Private Property Set Foo(ByVal arg2 As String)
End Property";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Param(s) to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[0].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveFirstParamFromGetterAndLetter()
        {
            //Input
            const string inputCode =
                @"Private Property Get Foo(ByVal arg1 As Integer)
End Property

Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Property";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Property Get Foo()
End Property

Private Property Let Foo(ByVal arg2 As String)
End Property";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Param(s) to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[0].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_SignatureContainsOptionalParam()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, Optional ByVal arg2 As String)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo arg1
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(Optional ByVal arg2 As String)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo 
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[0].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_RemoveOptionalParam()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, Optional ByVal arg2 As String)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo arg1
    Foo 1, ""test""
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal arg1 As Integer)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo arg1
    Foo 1
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[1].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_SignatureOnMultipleLines()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, _
                  ByVal arg2 As String, _
                  ByVal arg3 As Date)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal arg2 As String, _
                  ByVal arg3 As Date)
End Sub";   // note: VBE removes excess spaces

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[0].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_SignatureOnMultipleLines_RemoveSecond()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, _
                  ByVal arg2 As String, _
                  ByVal arg3 As Date)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal arg1 As Integer, _
                  ByVal arg3 As Date)
End Sub";   // note: VBE removes excess spaces

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[1].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_SignatureOnMultipleLines_RemoveLast()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, _
                  ByVal arg2 As String, _
                  ByVal arg3 As Date)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal arg1 As Integer, _
                  ByVal arg2 As String)
End Sub";   // note: VBE removes excess spaces

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[2].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_PassTargetIn()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, _
                  ByVal arg2 As String, _
                  ByVal arg3 As Date)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal arg2 As String, _
                  ByVal arg3 As Date)
End Sub";   // note: VBE removes excess spaces

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[0].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(model.TargetDeclaration);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void RemoveParametersRefactoring_CallOnMultipleLines()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date)
End Sub

Private Sub Goo(ByVal arg1 as Integer, ByVal arg2 As String, ByVal arg3 As Date)

    Foo arg1, _
        arg2, _
        arg3

End Sub
";
            var selection = new Selection(1, 16, 1, 16);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal arg2 As String, ByVal arg3 As Date)
End Sub

Private Sub Goo(ByVal arg1 as Integer, ByVal arg2 As String, ByVal arg3 As Date)

    Foo arg2, _
        arg3

End Sub
";   // note: IDE removes excess spaces

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, null);
                model.Parameters[0].IsRemoved = true;

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RemoveParametersRefactoring(vbe.Object, factory.Object);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
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
                model.Parameters[1].IsRemoved = true;

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
                model.Parameters[1].IsRemoved = true;

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
                model.Parameters[1].IsRemoved = true;

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
                model.Parameters[1].IsRemoved = true;

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
        public void ReorderParametersRefactoring_LastEventParamRemoved_EventImplementationSelected()
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
                model.Parameters.Last().IsRemoved = true;

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
                model.Parameters[1].IsRemoved = true;

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
                model.Parameters[1].IsRemoved = true;

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
                messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                    .Returns(DialogResult.Yes);

                //Specify Params to remove
                var model = new RemoveParametersModel(state, qualifiedSelection, messageBox.Object);
                model.Parameters[1].IsRemoved = true;

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
                messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                    .Returns(DialogResult.No);

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var model = new RemoveParametersModel(state, qualifiedSelection, new Mock<IMessageBox>().Object);
                model.Parameters[0].IsRemoved = true;

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

                var messageBox = new Mock<IMessageBox>();
                messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>())).Returns(DialogResult.OK);

                var factory = new RemoveParametersPresenterFactory(vbe.Object, null, state, messageBox.Object);
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
