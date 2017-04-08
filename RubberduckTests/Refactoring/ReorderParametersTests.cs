using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class ReorderParametersTests
    {
        [TestMethod]
        public void ReorderParams_SwapPositions()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //set up model
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParams_SwapPositions_SignatureContainsParamName()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(a, ba)
End Sub";
            var selection = new Selection(1, 16, 1, 16);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ba, a)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //set up model
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParams_SwapPositions_ReferenceValueContainsOtherReferenceValue()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(a, ba)
End Sub

Sub Goo()
    Foo 1, 121
End Sub";
            var selection = new Selection(1, 16, 1, 16);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ba, a)
End Sub

Sub Goo()
    Foo 121, 1
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //set up model
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParams_RefactorDeclaration()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //set up model
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(model.TargetDeclaration);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParams_RefactorDeclaration_FailsInvalidTarget()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //set up model
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);

            try
            {
                refactoring.Refactor(
                    model.Declarations.FirstOrDefault(
                        i => i.DeclarationType == Rubberduck.Parsing.Symbols.DeclarationType.ProceduralModule));
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Invalid declaration type", e.Message);
                return;
            }

            Assert.Fail();
        }

        [TestMethod]
        public void ReorderParams_WithOptionalParam()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String, Optional ByVal arg3 As Boolean = True)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer, Optional ByVal arg3 As Boolean = True)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //set up model
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[0],
                model.Parameters[2]
            };

            model.Parameters = reorderedParams;

            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParams_SwapPositions_UpdatesCallers()
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
@"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Sub

Private Sub Bar()
    Foo ""Hello"", 10
End Sub
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //set up model
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_ClientReferencesAreUpdated_ParensAroundCall()
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
    y = foo(42, x)
    Debug.Print y, x
End Sub

Private Function foo(ByVal b As Integer, ByRef a As Integer) As Integer
    a = b
    foo = a + b
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //set up model
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderNamedParams()
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
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg3 As Double, ByVal arg2 As String)
End Sub

Public Sub Goo()
    Foo arg2:=""test44"", arg1:=3, arg3:=6.1
End Sub
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[0],
                model.Parameters[2],
                model.Parameters[1]
            };

            model.Parameters = reorderedParams;

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderNamedParams_Function()
        {
            //Input
            const string inputCode =
@"Public Function Foo(ByVal arg1 As Integer, ByVal arg2 As String) As Boolean
    Foo = True
End Function";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Public Function Foo(ByVal arg2 As String, ByVal arg1 As Integer) As Boolean
    Foo = True
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderNamedParams_WithOptionalParam()
        {
            //Input
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String, Optional ByVal arg3 As Double)
End Sub

Public Sub Goo()
    Foo arg2:=""test44"", arg1:=3
End Sub
";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Public Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer, Optional ByVal arg3 As Double)
End Sub

Public Sub Goo()
    Foo arg1:=3, arg2:=""test44""
End Sub
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[0],
                model.Parameters[2]
            };

            model.Parameters = reorderedParams;

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderGetter()
        {
            //Input
            const string inputCode =
@"Private Property Get Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date) As Boolean
End Property";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Property Get Foo(ByVal arg2 As String, ByVal arg3 As Date, ByVal arg1 As Integer) As Boolean
End Property";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[2],
                model.Parameters[0]
            };

            model.Parameters = reorderedParams;

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderLetter()
        {
            //Input
            const string inputCode =
@"Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date)
End Property";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Property Let Foo(ByVal arg2 As String, ByVal arg1 As Integer, ByVal arg3 As Date)
End Property";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderSetter()
        {
            //Input
            const string inputCode =
@"Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date)
End Property";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Property Set Foo(ByVal arg2 As String, ByVal arg1 As Integer, ByVal arg3 As Date)
End Property";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderLastParamFromSetter_NotAllowed()
        {
            //Input
            const string inputCode =
@"Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new ReorderParametersModel(state, qualifiedSelection, null);

            Assert.AreEqual(1, model.Parameters.Count); // doesn't allow removing last param from setter
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderLastParamFromLetter_NotAllowed()
        {
            //Input
            const string inputCode =
@"Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new ReorderParametersModel(state, qualifiedSelection, null);

            Assert.AreEqual(1, model.Parameters.Count); // doesn't allow removing last param from letter
        }

        [TestMethod]
        public void ReorderParametersRefactoring_SignatureOnMultipleLines()
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
@"Private Sub Foo(ByVal arg3 As Date, _
                  ByVal arg2 As String, _
                  ByVal arg1 As Integer)
End Sub";   // note: IDE removes excess spaces

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_CallOnMultipleLines()
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
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg3 As Date, ByVal arg2 As String, ByVal arg1 As Integer)
End Sub

Private Sub Goo(ByVal arg1 as Integer, ByVal arg2 As String, ByVal arg3 As Date)

    Foo arg3, _
        arg2, _
        arg1

End Sub
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ClientReferencesAreNotUpdated_ParamArray()
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.OK, MessageBoxIcon.Warning)).Returns(DialogResult.OK);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(inputCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ClientReferencesAreUpdated_ParamArray()
        {
            //Input
            const string inputCode =
@"Sub Foo(ByVal arg1 As String, ByVal arg2 As Date, ParamArray arg3())
End Sub

Public Sub Goo(ByVal arg As Date, _
               ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
    Foo ""test"", arg, test1x, test2x, test3x, test4x, test5x, test6x
End Sub
";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Sub Foo(ByVal arg2 As Date, ByVal arg1 As String, ParamArray arg3())
End Sub

Public Sub Goo(ByVal arg As Date, _
               ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
    Foo arg, ""test"", test1x, test2x, test3x, test4x, test5x, test6x
End Sub
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>
            {
                model.Parameters[1],
                model.Parameters[0],
                model.Parameters[2]
            };

            model.Parameters = reorderedParams;

            //SetupFactory
            var factory = SetupFactory(model);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.OK, MessageBoxIcon.Warning)).Returns(DialogResult.OK);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ClientReferencesAreUpdated_ParamArray_CallOnMultiplelines()
        {
            //Input
            const string inputCode =
@"Sub Foo(ByVal arg1 As String, ByVal arg2 As Date, ParamArray arg3())
End Sub

Public Sub Goo(ByVal arg As Date, _
               ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
    Foo ""test"", _
        arg, _
        test1x, _
        test2x, _
        test3x, _
        test4x, _
        test5x, _
        test6x
End Sub
";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Sub Foo(ByVal arg2 As Date, ByVal arg1 As String, ParamArray arg3())
End Sub

Public Sub Goo(ByVal arg As Date, _
               ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
    Foo arg, _
        ""test"", _
        test1x, _
        test2x, _
        test3x, _
        test4x, _
        test5x, _
        test6x
End Sub
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[0],
                model.Parameters[2]
            };

            model.Parameters = reorderedParams;

            //SetupFactory
            var factory = SetupFactory(model);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.OK, MessageBoxIcon.Warning)).Returns(DialogResult.OK);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParams_MoveOptionalParamBeforeNonOptionalParamFails()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, Optional ByVal arg2 As String, Optional ByVal arg3 As Boolean = True)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //set up model
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[2],
                model.Parameters[0]
            };

            model.Parameters = reorderedParams;

            var factory = SetupFactory(model);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.OK, MessageBoxIcon.Warning)).Returns(DialogResult.OK);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(inputCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParams_ReorderCallsWithoutOptionalParams()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String, Optional ByVal arg3 As Boolean = True)
End Sub

Private Sub Goo(ByVal arg1 As Integer, ByVal arg2 As String)
    Foo arg1, arg2
End Sub
";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer, Optional ByVal arg3 As Boolean = True)
End Sub

Private Sub Goo(ByVal arg1 As Integer, ByVal arg2 As String)
    Foo arg2, arg1
End Sub
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //set up model
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[0],
                model.Parameters[2]
            };

            model.Parameters = reorderedParams;

            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderFirstParamFromGetterAndSetter()
        {
            //Input
            const string inputCode =
@"Private Property Get Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Property

Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date)
End Property";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Property Get Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Property

Private Property Set Foo(ByVal arg2 As String, ByVal arg1 As Integer, ByVal arg3 As Date)
End Property";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Param(s) to reorder
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderFirstParamFromGetterAndLetter()
        {
            //Input
            const string inputCode =
@"Private Property Get Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Property

Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date)
End Property";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Property Get Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Property

Private Property Let Foo(ByVal arg2 As String, ByVal arg1 As Integer, ByVal arg3 As Date)
End Property";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParams_PresenterIsNull()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var factory = new ReorderParametersPresenterFactory(vbe.Object, null, state, null);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory, null);
            refactoring.Refactor();

            Assert.AreEqual(inputCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_InterfaceParamsSwapped()
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
@"Public Sub DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";
            const string expectedCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";   // note: IDE removes excess spaces

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);
            var module1 = project.Object.VBComponents[0].CodeModule;
            var module2 = project.Object.VBComponents[1].CodeModule;

            //Specify Params to remove
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode1, module1.Content());
            Assert.AreEqual(expectedCode2, module2.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_InterfaceParamsSwapped_ParamsHaveDifferentNames()
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
@"Public Sub DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";
            const string expectedCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal v2 As String, ByVal v1 As Integer)
End Sub";   // note: IDE removes excess spaces

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);
            var module1 = project.Object.VBComponents[0].CodeModule;
            var module2 = project.Object.VBComponents[1].CodeModule;

            //Specify Params to remove
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode1, module1.Content());
            Assert.AreEqual(expectedCode2, module2.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_InterfaceParamsSwapped_ParamsHaveDifferentNames_TwoImplementations()
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
@"Public Sub DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";
            const string expectedCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal v2 As String, ByVal v1 As Integer)
End Sub";   // note: IDE removes excess spaces
            const string expectedCode3 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal s As String, ByVal i As Integer)
End Sub";   // note: IDE removes excess spaces

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode3)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);
            var module1 = project.Object.VBComponents[0].CodeModule;
            var module2 = project.Object.VBComponents[1].CodeModule;
            var module3 = project.Object.VBComponents[2].CodeModule;

            //Specify Params to remove
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode1, module1.Content());
            Assert.AreEqual(expectedCode2, module2.Content());
            Assert.AreEqual(expectedCode3, module3.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_InterfaceParamsSwapped_AcceptPrompt()
        {
            //Input
            const string inputCode1 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
@"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";

            var selection = new Selection(3, 23, 3, 27);

            //Expectation
            const string expectedCode1 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";   // note: IDE removes excess spaces

            const string expectedCode2 =
@"Public Sub DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);
            var module1 = project.Object.VBComponents[0].CodeModule;
            var module2 = project.Object.VBComponents[1].CodeModule;

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(
                m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                .Returns(DialogResult.Yes);

            //Specify Params to remove
            var model = new ReorderParametersModel(state, qualifiedSelection, messageBox.Object);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode1, module1.Content());
            Assert.AreEqual(expectedCode2, module2.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ParamsSwapped_RejectPrompt()
        {
            //Input
            const string inputCode1 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
@"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";

            var selection = new Selection(3, 23, 3, 27);

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.No);

            //Specify Params to remove
            var model = new ReorderParametersModel(state, qualifiedSelection, messageBox.Object);
            Assert.IsNull(model.TargetDeclaration);
        }

        [TestMethod]
        public void ReorderParametersRefactoring_EventParamsSwapped()
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
@"Public Event Foo(ByVal arg2 As String, ByVal arg1 As Integer)";

            const string expectedCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Sub";   // note: IDE removes excess spaces

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);
            var module1 = project.Object.VBComponents[0].CodeModule;
            var module2 = project.Object.VBComponents[1].CodeModule;

            //Specify Params to remove
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode1, module1.Content());
            Assert.AreEqual(expectedCode2, module2.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_EventParamsSwapped_EventImplementationSelected()
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

Private Sub abc_Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Sub";   // note: IDE removes excess spaces

            const string expectedCode2 =
@"Public Event Foo(ByVal arg2 As String, ByVal arg1 As Integer)";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);
            var module1 = project.Object.VBComponents[0].CodeModule;
            var module2 = project.Object.VBComponents[1].CodeModule;

            //Specify Params to remove
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode1, module1.Content());
            Assert.AreEqual(expectedCode2, module2.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_EventParamsSwapped_DifferentParamNames()
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
@"Public Event Foo(ByVal arg2 As String, ByVal arg1 As Integer)";

            const string expectedCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal s As String, ByVal i As Integer)
End Sub";   // note: IDE removes excess spaces

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);
            var module1 = project.Object.VBComponents[0].CodeModule;
            var module2 = project.Object.VBComponents[1].CodeModule;

            //Specify Params to remove
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode1, module1.Content());
            Assert.AreEqual(expectedCode2, module2.Content());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_EventParamsSwapped_DifferentParamNames_TwoHandlers()
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
@"Public Event Foo(ByVal arg2 As String, ByVal arg1 As Integer)";

            const string expectedCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal s As String, ByVal i As Integer)
End Sub";   // note: IDE removes excess spaces

            const string expectedCode3 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal v2 As String, ByVal v1 As Integer)
End Sub";   // note: IDE removes excess spaces

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class3", ComponentType.ClassModule, inputCode3)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);
            var module1 = project.Object.VBComponents[0].CodeModule;
            var module2 = project.Object.VBComponents[1].CodeModule;
            var module3 = project.Object.VBComponents[2].CodeModule;

            //Specify Params to remove
            var model = new ReorderParametersModel(state, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ReorderParametersRefactoring(vbe.Object, factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode1, module1.Content());
            Assert.AreEqual(expectedCode2, module2.Content());
            Assert.AreEqual(expectedCode3, module3.Content());
        }

        [TestMethod]
        public void Presenter_ParameterlessMemberCreatesNullModel()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.OK);

            var factory = new ReorderParametersPresenterFactory(vbe.Object, null, state, messageBox.Object);
            var presenter = factory.Create();

            var result = presenter.Show();

            Assert.IsNull(result);
        }

        [TestMethod]
        public void Presenter_SingleParameterMemberCreatesNullModel()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer)
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.OK);

            var factory = new ReorderParametersPresenterFactory(vbe.Object, null, state, messageBox.Object);
            var presenter = factory.Create();

            var result = presenter.Show();

            Assert.IsNull(result);
        }

        [TestMethod]
        public void Presenter_NullTargetCreatesNullModel()
        {
            //Input
            const string inputCode =
@"
Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 1, 1, 1);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var factory = new ReorderParametersPresenterFactory(vbe.Object, null, state, null);

            var presenter = factory.Create();

            var result = presenter.Show();

            Assert.IsNull(result);
        }

        [TestMethod]
        public void Factory_NullSelectionCreatesNullPresenter()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            vbe.Setup(v => v.ActiveCodePane).Returns((ICodePane)null);

            var factory = new ReorderParametersPresenterFactory(vbe.Object, null, state, null);

            var result = factory.Create();

            Assert.IsNull(result);
        }

        #region setup
        private static Mock<IRefactoringPresenterFactory<IReorderParametersPresenter>> SetupFactory(ReorderParametersModel model)
        {
            var presenter = new Mock<IReorderParametersPresenter>();
            presenter.Setup(p => p.Show()).Returns(model);

            var factory = new Mock<IRefactoringPresenterFactory<IReorderParametersPresenter>>();
            factory.Setup(f => f.Create()).Returns(presenter.Object);
            return factory;
        }

        #endregion
    }
}
