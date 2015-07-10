using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class ReorderParametersTests : VbeTestBase
    {
        [TestMethod]
        public void ReorderParams_SwapPositions()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Sub";

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //set up model
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ReorderParams_RefactorDeclaration()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Sub";

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //set up model
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(model.TargetDeclaration);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ReorderParams_RefactorDeclaration_FailsInvalidTarget()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Arrange
            var project = SetupMockProject(inputCode);
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //set up model
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);

            //assert
            try
            {
                refactoring.Refactor(
                    model.Declarations.Items.FirstOrDefault(
                        i => i.DeclarationType == Rubberduck.Parsing.Symbols.DeclarationType.Class));
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Invalid declaration type", e.Message);
                return;
            }

            Assert.IsTrue(false);
        }

        [TestMethod]
        public void ReorderParams_WithOptionalParam()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String, Optional ByVal arg3 As Boolean = True)
End Sub";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer, Optional ByVal arg3 As Boolean = True)
End Sub";

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //set up model
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[0],
                model.Parameters[2]
            };

            model.Parameters = reorderedParams;

            var factory = SetupFactory(model);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Sub

Private Sub Bar()
 Foo ""Hello"", 10
End Sub
";

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //set up model
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg3 As Double, ByVal arg2 As String)
End Sub

Public Sub Goo()
 Foo arg2:=""test44"", arg1:=3, arg3:=6.1
End Sub
";

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[0],
                model.Parameters[2],
                model.Parameters[1]
            };

            model.Parameters = reorderedParams;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderNamedParams_Function()
        {
            //Input
            const string inputCode =
@"Public Function Foo(ByVal arg1 As Integer, ByVal arg2 As String) As Boolean
    Foo = True
End Function";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Public Function Foo(ByVal arg2 As String, ByVal arg1 As Integer) As Boolean
    Foo = True
End Function";

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Public Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer, Optional ByVal arg3 As Double)
End Sub

Public Sub Goo()
 Foo arg1:=3, arg2:=""test44""
End Sub
";

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[0],
                model.Parameters[2]
            };

            model.Parameters = reorderedParams;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderGetter()
        {
            //Input
            const string inputCode =
@"Private Property Get Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date) As Boolean
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Get Foo(ByVal arg2 As String, ByVal arg3 As Date, ByVal arg1 As Integer) As Boolean
End Property";

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[2],
                model.Parameters[0]
            };

            model.Parameters = reorderedParams;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderLetter()
        {
            //Input
            const string inputCode =
@"Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date) 
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Let Foo(ByVal arg2 As String, ByVal arg1 As Integer, ByVal arg3 As Date)
End Property";

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderSetter()
        {
            //Input
            const string inputCode =
@"Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date) 
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Set Foo(ByVal arg2 As String, ByVal arg1 As Integer, ByVal arg3 As Date)
End Property";

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderLastParamFromSetter_NotAllowed()
        {
            //Input
            const string inputCode =
@"Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Arrange
            var project = SetupMockProject(inputCode);
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);

            // Assert
            Assert.AreEqual(1, model.Parameters.Count); // doesn't allow removing last param from setter
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderLastParamFromLetter_NotAllowed()
        {
            //Input
            const string inputCode =
@"Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Arrange
            var project = SetupMockProject(inputCode);
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);

            // Assert
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
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg3 As Date,                  ByVal arg2 As String,                  ByVal arg1 As Integer)


End Sub";   // note: IDE removes excess spaces

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg3 As Date, ByVal arg2 As String, ByVal arg1 As Integer)
End Sub

Private Sub Goo(ByVal arg1 as Integer, ByVal arg2 As String, ByVal arg3 As Date)

 Foo arg3, arg2, arg1



End Sub
";

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.OK, MessageBoxIcon.Warning)).Returns(DialogResult.OK);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(inputCode, module.Lines());
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
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

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

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
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

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

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

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
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

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ReorderParams_MoveOptionalParamBeforeNonOptionalParamFails()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, Optional ByVal arg2 As String, Optional ByVal arg3 As Boolean = True)
End Sub";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //set up model
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
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

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(inputCode, module.Lines());
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
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer, Optional ByVal arg3 As Boolean = True)
End Sub

Private Sub Goo(ByVal arg1 As Integer, ByVal arg2 As String)
 Foo arg2, arg1
End Sub
";

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //set up model
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[0],
                model.Parameters[2]
            };

            model.Parameters = reorderedParams;

            var factory = SetupFactory(model);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Get Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Property

Private Property Set Foo(ByVal arg2 As String, ByVal arg1 As Integer, ByVal arg3 As Date)
End Property";

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Param(s) to reorder
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Get Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Property

Private Property Let Foo(ByVal arg2 As String, ByVal arg1 As Integer, ByVal arg3 As Date)
End Property";

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ReorderParams_PresenterIsNull()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";

            //Arrange
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns((QualifiedSelection?)null);

            var factory = new ReorderParametersPresenterFactory(editor.Object, null,
                parseResult, null);

            //act
            var refactoring = new ReorderParametersRefactoring(factory, null);
            refactoring.Refactor();

            Assert.AreEqual(inputCode, module.Lines());
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

            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode1 =
@"Public Sub DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";
            const string expectedCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";   // note: IDE removes excess spaces

            //Arrange
            var component1 = CreateMockComponent(inputCode1, "IClass1",
                vbext_ComponentType.vbext_ct_ClassModule);
            var component2 = CreateMockComponent(inputCode2, "Class1",
                vbext_ComponentType.vbext_ct_ClassModule);

            var project = CreateMockProject("VBEProject", vbext_ProjectProtection.vbext_pp_none,
                new List<Mock<VBComponent>>() { component1, component2 });
            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
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

            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode1 =
@"Public Sub DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";
            const string expectedCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal v2 As String, ByVal v1 As Integer)
End Sub";   // note: IDE removes excess spaces

            //Arrange
            var component1 = CreateMockComponent(inputCode1, "IClass1",
                vbext_ComponentType.vbext_ct_ClassModule);
            var component2 = CreateMockComponent(inputCode2, "Class1",
                vbext_ComponentType.vbext_ct_ClassModule);

            var project = CreateMockProject("VBEProject", vbext_ProjectProtection.vbext_pp_none,
                new List<Mock<VBComponent>>() { component1, component2 });
            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
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

            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

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

            //Arrange
            var component1 = CreateMockComponent(inputCode1, "IClass1",
                vbext_ComponentType.vbext_ct_ClassModule);
            var component2 = CreateMockComponent(inputCode2, "Class1",
                vbext_ComponentType.vbext_ct_ClassModule);
            var component3 = CreateMockComponent(inputCode3, "Class2",
                vbext_ComponentType.vbext_ct_ClassModule);

            var project = CreateMockProject("VBEProject", vbext_ProjectProtection.vbext_pp_none,
                new List<Mock<VBComponent>>() { component1, component2, component3 });
            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;
            var module3 = project.Object.VBComponents.Item(2).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
            Assert.AreEqual(expectedCode3, module3.Lines());
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

            var selection = new Selection(3, 23, 3, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode1 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";   // note: IDE removes excess spaces

            const string expectedCode2 =
@"Public Sub DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";

            //Arrange
            var component1 = CreateMockComponent(inputCode1, "Class1",
                vbext_ComponentType.vbext_ct_ClassModule);
            var component2 = CreateMockComponent(inputCode2, "IClass1",
                vbext_ComponentType.vbext_ct_ClassModule);

            var project = CreateMockProject("VBEProject", vbext_ProjectProtection.vbext_pp_none,
                new List<Mock<VBComponent>>() { component1, component2 });
            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(
                m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                .Returns(DialogResult.Yes);

            //Specify Params to remove
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, messageBox.Object);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
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

            var selection = new Selection(3, 23, 3, 27); //startLine, startCol, endLine, endCol

            //Arrange
            var component1 = CreateMockComponent(inputCode1, "Class1",
                vbext_ComponentType.vbext_ct_ClassModule);
            var component2 = CreateMockComponent(inputCode2, "IClass1",
                vbext_ComponentType.vbext_ct_ClassModule);

            var project = CreateMockProject("VBEProject", vbext_ProjectProtection.vbext_pp_none,
                new List<Mock<VBComponent>>() { component1, component2 });
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(
                m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                .Returns(DialogResult.No);

            //Specify Params to remove
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, messageBox.Object);
            Assert.AreEqual(null, model.TargetDeclaration);
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

            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode1 =
@"Public Event Foo(ByVal arg2 As String, ByVal arg1 As Integer)";

            const string expectedCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Sub";   // note: IDE removes excess spaces

            //Arrange
            var component1 = CreateMockComponent(inputCode1, "Class1",
                vbext_ComponentType.vbext_ct_ClassModule);
            var component2 = CreateMockComponent(inputCode2, "Class2",
                vbext_ComponentType.vbext_ct_ClassModule);

            var project = CreateMockProject("VBEProject", vbext_ProjectProtection.vbext_pp_none,
                new List<Mock<VBComponent>>() { component1, component2 });
            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
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

            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode1 =
@"Public Event Foo(ByVal arg2 As String, ByVal arg1 As Integer)";

            const string expectedCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal s As String, ByVal i As Integer)
End Sub";   // note: IDE removes excess spaces

            //Arrange
            var component1 = CreateMockComponent(inputCode1, "Class1",
                vbext_ComponentType.vbext_ct_ClassModule);
            var component2 = CreateMockComponent(inputCode2, "Class2",
                vbext_ComponentType.vbext_ct_ClassModule);

            var project = CreateMockProject("VBEProject", vbext_ProjectProtection.vbext_pp_none,
                new List<Mock<VBComponent>>() { component1, component2 });
            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
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

            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

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

            //Arrange
            var component1 = CreateMockComponent(inputCode1, "Class1",
                vbext_ComponentType.vbext_ct_ClassModule);
            var component2 = CreateMockComponent(inputCode2, "Class2",
                vbext_ComponentType.vbext_ct_ClassModule);
            var component3 = CreateMockComponent(inputCode3, "Class3",
                vbext_ComponentType.vbext_ct_ClassModule);

            var project = CreateMockProject("VBEProject", vbext_ProjectProtection.vbext_pp_none,
                new List<Mock<VBComponent>>() { component1, component2, component3 });
            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;
            var module3 = project.Object.VBComponents.Item(2).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new ReorderParametersModel(parseResult, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
            Assert.AreEqual(expectedCode3, module3.Lines());
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithParametersChanged()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Arrange
            var project = SetupMockProject(inputCode);
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var model = new ReorderParametersModel(parseResult, qualifiedSelection, new RubberduckMessageBox());
            model.Parameters.Reverse();

            var view = new Mock<IReorderParametersView>();
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);
            view.Setup(v => v.Parameters).Returns(model.Parameters);

            var factory = new ReorderParametersPresenterFactory(editor.Object, view.Object,
                parseResult, null);

            var presenter = factory.Create();

            Assert.AreEqual(model.Parameters, presenter.Show().Parameters);
        }

        [TestMethod]
        public void Presenter_Reject_ReturnsNull()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Arrange
            var project = SetupMockProject(inputCode);
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var model = new ReorderParametersModel(parseResult, qualifiedSelection, new RubberduckMessageBox());

            var view = new Mock<IReorderParametersView>();
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.Cancel);
            view.Setup(v => v.Parameters).Returns(model.Parameters);

            var factory = new ReorderParametersPresenterFactory(editor.Object, view.Object,
                parseResult, null);

            var presenter = factory.Create();

            Assert.AreEqual(null, presenter.Show());
        }
        
        [TestMethod]
        public void Presenter_NoParams()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";
            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Arrange
            var project = SetupMockProject(inputCode);
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>())).Returns(DialogResult.OK);

            var factory = new ReorderParametersPresenterFactory(editor.Object, null,
                parseResult, messageBox.Object);

            var presenter = factory.Create();

            Assert.AreEqual(null, presenter.Show());
        }

        [TestMethod]
        public void Presenter_OneParam()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer)
End Sub";
            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Arrange
            var project = SetupMockProject(inputCode);
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>())).Returns(DialogResult.OK);

            var factory = new ReorderParametersPresenterFactory(editor.Object, null,
                parseResult, messageBox.Object);

            var presenter = factory.Create();

            Assert.AreEqual(null, presenter.Show());
        }

        [TestMethod]
        public void Presenter_TargetIsNull()
        {
            //Input
            const string inputCode =
@"
Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 1, 1, 1); //startLine, startCol, endLine, endCol

            //Arrange
            var project = SetupMockProject(inputCode);
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var factory = new ReorderParametersPresenterFactory(editor.Object, null,
                parseResult, null);

            var presenter = factory.Create();

            Assert.AreEqual(null, presenter.Show());
        }

        [TestMethod]
        public void Factory_SelectionIsNull()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";

            //Arrange
            var project = SetupMockProject(inputCode);
            var parseResult = new RubberduckParser().Parse(project.Object);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns((QualifiedSelection?)null);

            var factory = new ReorderParametersPresenterFactory(editor.Object, null,
                parseResult, null);

            Assert.AreEqual(null, factory.Create());
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
