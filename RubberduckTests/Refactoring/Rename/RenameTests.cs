using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.Rename;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.Refactoring.Rename
{
    [TestFixture]
    public class RenameTests : InteractiveRefactoringTestBase<IRenamePresenter, RenameModel>
    {
        internal const char FAUX_CURSOR = '|';

        #region Rename Variable Tests

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameVariable()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "val1", newName: "val2");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Private Sub Foo()
    Dim va|l1 As Integer
End Sub",
                Expected =
                    @"Private Sub Foo()
    Dim val2 As Integer
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameVariable_UpdatesReferences()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "val1", newName: "val2");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Private Sub Foo()
    Dim v|al1 As Integer
    val1 = val1 + 5
End Sub",
                Expected =
                    @"Private Sub Foo()
    Dim val2 As Integer
    val2 = val2 + 5
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        //See issue #5277 at https://github.com/rubberduck-vba/Rubberduck/issues/5277
        public void RenameRefactoring_RenameArray_FromDeclaration()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "arr", newName: "bar");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Private Sub Foo()
    Dim a|rr(0 To 1) As Integer
    arr(1) = arr(0)
End Sub",
                Expected =
                    @"Private Sub Foo()
    Dim bar(0 To 1) As Integer
    bar(1) = bar(0)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        //See issue #5277 at https://github.com/rubberduck-vba/Rubberduck/issues/5277
        public void RenameRefactoring_RenameReDimDeclaredArray_FromDeclaration()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "arr", newName: "bar");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Private Sub Foo()
    ReDim a|rr(0 To 1)
    arr(1) = arr(0)
End Sub",
                Expected =
                    @"Private Sub Foo()
    ReDim bar(0 To 1)
    bar(1) = bar(0)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        //See issue #5277 at https://github.com/rubberduck-vba/Rubberduck/issues/5277
        public void RenameRefactoring_RenameArray_FromReference()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "arr", newName: "bar");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Private Sub Foo()
    Dim arr(0 To 1) As Integer
    arr(1) = ar|r(0)
End Sub",
                Expected =
                    @"Private Sub Foo()
    Dim bar(0 To 1) As Integer
    bar(1) = bar(0)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        //See issue #5277 at https://github.com/rubberduck-vba/Rubberduck/issues/5277
        public void RenameRefactoring_RenameReDimDeclaredArray_FromReference()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "arr", newName: "bar");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Private Sub Foo()
    ReDim arr(0 To 1)
    arr(1) = a|rr(0)
End Sub",
                Expected =
                    @"Private Sub Foo()
    ReDim bar(0 To 1)
    bar(1) = bar(0)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        //See issue #5236 at https://github.com/rubberduck-vba/Rubberduck/issues/5236
        public void RenameRefactoring_RenameForIndex_UpdatesReferences()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "loopIndex", newName: "otherLoopIndex");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Private Sub Foo()
    Dim loop|Index As Long
    For loopIndex = 0 To 42
        'DoSomething
    Next loopIndex
End Sub",
                Expected =
                    @"Private Sub Foo()
    Dim otherLoopIndex As Long
    For otherLoopIndex = 0 To 42
        'DoSomething
    Next otherLoopIndex
End Sub",
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        #endregion
        #region Rename Parameter Tests

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameParameter()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "arg1", newName: "arg2");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Private Sub Foo(ByVal ar|g1 As String)
End Sub",
                Expected =
                    @"Private Sub Foo(ByVal arg2 As String)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameMulitlinedParameter()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "arg3", newName: "arg2");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Private Sub Foo(ByVal arg1 As String, _
        ByVal ar|g3 As String)
End Sub",
                Expected =
                    @"Private Sub Foo(ByVal arg1 As String, _
        ByVal arg2 As String)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameParameter_UpdatesReferences()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "arg1", newName: "arg2");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Private Sub Foo(ByVal ar|g1 As String)
    arg1 = ""test""
End Sub",
                Expected =
                    @"Private Sub Foo(ByVal arg2 As String)
    arg2 = ""test""
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameFirstPropertyParameter_UpdatesAllRelatedParameters()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "index", newName: "renamed");
            var inputOutput = new RenameTestModuleDefinition("ClassFoo")
            {
                Input =
                    @"Property Get Foo(ByVal in|dex As Integer) As Variant
    Dim d As Integer
    d = index
End Property

Property Let Foo(ByVal index As Integer, ByVal value As Variant)
    Dim d As Integer
    d = index
End Property

Property Set Foo(ByVal index As Integer, ByVal value As Variant)
    Dim d As Integer
    d = index
End Property",

                Expected =
                    @"Property Get Foo(ByVal renamed As Integer) As Variant
    Dim d As Integer
    d = renamed
End Property

Property Let Foo(ByVal renamed As Integer, ByVal value As Variant)
    Dim d As Integer
    d = renamed
End Property

Property Set Foo(ByVal renamed As Integer, ByVal value As Variant)
    Dim d As Integer
    d = renamed
End Property"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameFirstPropertyParameter_DoesNotUpdateUnrelatedParameters()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "index", newName: "renamed");
            var inputOutput = new RenameTestModuleDefinition("ClassFoo")
            {
                Input =
                    @"Property Get Foo(ByVal in|dex As Integer) As Variant
    Dim d As Integer
    d = index
End Property

Property Let Foo(ByVal index As Integer, ByVal value As Variant)
    Dim d As Integer
    d = index
End Property

Property Set Bar(ByVal index As Integer, ByVal value As Variant)
    Dim d As Integer
    d = index
End Property",

                Expected =
                    @"Property Get Foo(ByVal renamed As Integer) As Variant
    Dim d As Integer
    d = renamed
End Property

Property Let Foo(ByVal renamed As Integer, ByVal value As Variant)
    Dim d As Integer
    d = renamed
End Property

Property Set Bar(ByVal index As Integer, ByVal value As Variant)
    Dim d As Integer
    d = index
End Property"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameFirstPropertyParameter_DoesNotUpdateOtherModules()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "index", newName: "renamed");
            var inputOutput = new RenameTestModuleDefinition("ClassFoo")
            {
                Input =
                    @"Property Get Foo(ByVal in|dex As Integer) As Variant
    Dim d As Integer
    d = index
End Property

Property Let Foo(ByVal index As Integer, ByVal value As Variant)
    Dim d As Integer
    d = index
End Property",

                Expected =
                    @"Property Get Foo(ByVal renamed As Integer) As Variant
    Dim d As Integer
    d = renamed
End Property

Property Let Foo(ByVal renamed As Integer, ByVal value As Variant)
    Dim d As Integer
    d = renamed
End Property"
            };

            var secondInputOutput = new RenameTestModuleDefinition("ClassBar")
            {
                Input =
                    @"Property Get Foo(ByVal index As Integer) As Variant
    Dim d As Integer
    d = index
End Property

Property Let Foo(ByVal index As Integer, ByVal value As Variant)
    Dim d As Integer
    d = index
End Property",

                Expected =
                    @"Property Get Foo(ByVal index As Integer) As Variant
    Dim d As Integer
    d = index
End Property

Property Let Foo(ByVal index As Integer, ByVal value As Variant)
    Dim d As Integer
    d = index
End Property"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput, secondInputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameLastPropertyParameter_UpdatesAllRelatedParameters()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "value", newName: "renamed");
            var inputOutput = new RenameTestModuleDefinition("ClassFoo")
            {
                Input =
                    @"Property Let Foo(ByVal index As Integer, ByVal va|lue As Variant)
    Dim d As Variant
    d = value
End Property

Property Set Foo(ByVal index As Integer, ByVal value As Variant)
    Dim d As Variant
    d = value
End Property",
                Expected =
                    @"Property Let Foo(ByVal index As Integer, ByVal renamed As Variant)
    Dim d As Variant
    d = renamed
End Property

Property Set Foo(ByVal index As Integer, ByVal renamed As Variant)
    Dim d As Variant
    d = renamed
End Property"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameLastPropertyParameter_DoesNotUpdateUnrelatedParameters()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "value", newName: "renamed");
            var inputOutput = new RenameTestModuleDefinition("ClassFoo")
            {
                Input =
                    @"Property Let Foo(ByVal index As Integer, ByVal va|lue As Variant)
    Dim d As Variant
    d = value
End Property

Property Set Bar(ByVal index As Integer, ByVal value As Variant)
    Dim d As Variant
    d = value
End Property",
                Expected =
                    @"Property Let Foo(ByVal index As Integer, ByVal renamed As Variant)
    Dim d As Variant
    d = renamed
End Property

Property Set Bar(ByVal index As Integer, ByVal value As Variant)
    Dim d As Variant
    d = value
End Property"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameLastPropertyParameter_UpdatesRelatedParametersWithSameName()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "value", newName: "renamed");
            var inputOutput = new RenameTestModuleDefinition("ClassFoo")
            {
                Input =
                    @"Property Get Foo(ByVal index As Integer) As Variant
End Property

Property Let Foo(ByVal index As Integer, ByVal v|alue As Variant)
    Dim d As Variant
    d = value
End Property

Property Set Foo(ByVal index As Integer, ByVal fizz As Variant)
    Dim d As Variant
    d = fizz
End Property",
                Expected =
                    @"Property Get Foo(ByVal index As Integer) As Variant
End Property

Property Let Foo(ByVal index As Integer, ByVal renamed As Variant)
    Dim d As Variant
    d = renamed
End Property

Property Set Foo(ByVal index As Integer, ByVal fizz As Variant)
    Dim d As Variant
    d = fizz
End Property"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        #endregion
        #region Rename Member Tests

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameSub_WarnConflictingName_CancelIfNotAccepted()
        {
            var moduleCode =
                @"Private Sub Fo|o()
    Dim Goo As Integer
End Sub";

            Assert_WarnsAboutRenameConflict_ConfirmationOutcomeHasExpectedResult(
                moduleCode,
                ComponentType.StandardModule,
                "Foo",
                DeclarationType.Procedure,
                "Goo",
                false,
                RefactoringDialogResult.Cancel);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameSub_WarnConflictingName_ProceedIfAccepted()
        {
            var moduleCode =
                @"Private Sub Fo|o()
    Dim Goo As Integer
End Sub";

            Assert_WarnsAboutRenameConflict_ConfirmationOutcomeHasExpectedResult(
                moduleCode,
                ComponentType.StandardModule,
                "Foo",
                DeclarationType.Procedure,
                "Goo",
                true,
                RefactoringDialogResult.Execute);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameSub_ConflictingNames_Accept()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "Foo", newName: "Goo");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Private Sub Fo|o()
    Dim Goo As Integer
End Sub",
                Expected =
                    @"Private Sub Goo()
    Dim Goo As Integer
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameSub_UpdatesReferences()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "Foo", newName: "Hoo");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Private Sub Fo|o()
End Sub

Private Sub Goo()
    Foo
End Sub",
                Expected =
                    @"Private Sub Hoo()
End Sub

Private Sub Goo()
    Hoo
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }
        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameGetterAndSetter()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "Foo", newName: "Goo");
            var inputOutput = new RenameTestModuleDefinition("ClassFoo")
            {
                Input =
                    @"Private Property Get F|oo(ByVal arg1 As Integer) As String
    Foo = ""Hello""
End Property

Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property",
                Expected =
                    @"Private Property Get Goo(ByVal arg1 As Integer) As String
    Goo = ""Hello""
End Property

Private Property Set Goo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameGetterAndLetter()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "Foo", newName: "Goo");
            var inputOutput = new RenameTestModuleDefinition("ClassFoo")
            {
                Input =
                    @"Private Property Get Foo() 
End Property

Private Property Let F|oo(ByVal arg1 As String) 
End Property",
                Expected =
                    @"Private Property Get Goo() 
End Property

Private Property Let Goo(ByVal arg1 As String) 
End Property"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameFunction()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "Foo", newName: "Hoo");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Private Function Foo() As Boolean
    Fo|o = True
End Function",
                Expected =
                    @"Private Function Hoo() As Boolean
    Hoo = True
End Function"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameFunction_UpdatesReferences()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "Foo", newName: "Hoo");
            var inputOutput = new RenameTestModuleDefinition("ClassFoo")
            {
                Input =
                    @"Private Function Fo|o() As Boolean
    Foo = True
End Function
Private Sub Goo()
    Dim var1 As Boolean
    var1 = Foo()
End Sub",
                Expected =
                    @"Private Function Hoo() As Boolean
    Hoo = True
End Function
Private Sub Goo()
    Dim var1 As Boolean
    var1 = Hoo()
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameFunction_DoesNotChangeIndexedDefaultMemberCalls()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "Foo", newName: "Goo");
            var defaultMemberInputOutput = new RenameTestModuleDefinition("ClassFoo")
            {
                Input =
                    @"Public Function Fo|o(arg As String) As Boolean
Attribute Foo.VB_UserMemId = 0
    Foo = True
End Function",
                //TODO: Make it possible that the attribute survives this appropriately adjusted.
                //The VBE will remove the now invalid attribute, which does not happen in tests.
                Expected =
                    @"Public Function Goo(arg As String) As Boolean
Attribute Foo.VB_UserMemId = 0
    Goo = True
End Function"
            };
            var callingModuleInputOutput = new RenameTestModuleDefinition("TestModule", ComponentType.StandardModule)
            {
                Input =
                    @"Private Function Baz(arg As String) As Boolean
    Dim bar As ClassFoo
    Set bar = New ClassFoo
    Baz = bar(arg)
End Function",
                Expected =
                    @"Private Function Baz(arg As String) As Boolean
    Dim bar As ClassFoo
    Set bar = New ClassFoo
    Baz = bar(arg)
End Function"
            };
            PerformExpectedVersusActualRenameTests(tdo, defaultMemberInputOutput, callingModuleInputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameFunction_DoesNotChangeNonIndexedDefaultMemberCalls()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "Foo", newName: "Goo");
            var defaultMemberInputOutput = new RenameTestModuleDefinition("ClassFoo")
            {
                Input =
                    @"Public Function Fo|o() As Boolean
Attribute Foo.VB_UserMemId = 0
    Foo = True
End Function",
                //TODO: Make it possible that the attribute survives this appropriately adjusted.
                //The VBE will remove the now invalid attribute, which does not happen in tests.
                Expected =
                    @"Public Function Goo() As Boolean
Attribute Foo.VB_UserMemId = 0
    Goo = True
End Function"
            };
            var callingModuleInputOutput = new RenameTestModuleDefinition("TestModule", ComponentType.StandardModule)
            {
                Input =
                    @"Private Function Baz(arg As String) As Boolean
    Dim bar As ClassFoo
    Set bar = New ClassFoo
    Baz = bar
End Function",
                Expected =
                    @"Private Function Baz(arg As String) As Boolean
    Dim bar As ClassFoo
    Set bar = New ClassFoo
    Baz = bar
End Function"
            };
            PerformExpectedVersusActualRenameTests(tdo, defaultMemberInputOutput, callingModuleInputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameVariableWithBracketedExpressionInModule()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "Foo", newName: "Hoo");
            var inputOutput = new RenameTestModuleDefinition("TestModule1", ComponentType.Document)
            {
                Input =
                    @"Private Fo|o() As Long

Public Sub Derp()
  [Something].Clear
End Sub",
                Expected =
                    @"Private Hoo() As Long

Public Sub Derp()
  [Something].Clear
End Sub"
            };

            tdo.UseLibraries = true;
            tdo.AdditionalSetup = t =>
            {
                var hostApp = new Mock<IHostApplication>();
                hostApp.Setup(x => x.ApplicationName).Returns("EXCEL");
                var mock = Mock.Get(tdo.VBE);
                mock.Setup(x => x.HostApplication()).Returns(hostApp.Object);
            };

            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        #endregion
        #region Rename Control Tests
        //All RenameControl tests are ignored because control renames depend on access to
        //Non-UserDefined declarations in the DeclarationFinder.  So, the control rename scenarios
        //below can only be tested if implemented (and tested) within Excel.  

        [Test, Ignore("")]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameControlFromEventHandler()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "cmdBtn1", newName: "cmdBigButton");
            var inputOutput = new RenameTestModuleDefinition("UserForm1", ComponentType.UserForm)
            {
                Input =
                    @"Private Sub cmdBtn1_Cl|ick()
End Sub

Private Sub tbEnterName_Change()
    cmdBtn1_Click 'bad idea, but someone will do it
End Sub

Private Sub UserForm_Click()
    cmdBtn1.Caption = ""Click This""
End Sub",
                Expected =
                    @"Private Sub cmdBigButton_Click()
End Sub

Private Sub tbEnterName_Change()
    cmdBigButton_Click 'bad idea, but someone will do it
End Sub

Private Sub UserForm_Click()
    cmdBigButton.Caption = ""Click This""
End Sub"
            };
            inputOutput.ControlNames.Add("cmdBtn1");
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test, Ignore("")]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameControlFromEventHandlerNameCollision()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "cmdBtn1", newName: "cmdBigButton");
            var inputOutput = new RenameTestModuleDefinition("UserForm1", ComponentType.UserForm)
            {
                Input =
                    @"Private Sub cmdBtn1_Cl|ick()
    cmdBtn1_PoorlyNamedHelper
End Sub

Private Sub tbEnterName_Change()
    cmdBtn1_Click 'bad idea, but someone will do it
End Sub

Private Sub cmdBtn1_PoorlyNamedHelper()
    cmdBtn1.Caption = ""Click This""
End Sub",
                Expected =
                    @"Private Sub cmdBigButton_Click()
    cmdBtn1_PoorlyNamedHelper
End Sub

Private Sub tbEnterName_Change()
    cmdBigButton_Click 'bad idea, but someone will do it
End Sub

Private Sub cmdBtn1_PoorlyNamedHelper()
    cmdBigButton.Caption = ""Click This""
End Sub"
            };
            inputOutput.ControlNames.Add("cmdBtn1");
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test, Ignore("")]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameControlRenameInReference()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "cmdBtn1", newName: "cmdBigButton");
            var inputOutput = new RenameTestModuleDefinition("UserForm1", ComponentType.UserForm)
            {
                Input =
                    @"Private Sub cmdBtn1_Click()
End Sub

Private Sub tbEnterName_Change()
    cmdBtn1_Click 'bad idea, but someone will do it
End Sub

Private Sub UserForm_Click()
    cmd|Btn1.Caption = ""Click This""
End Sub",
                Expected =
                    @"Private Sub cmdBigButton_Click()
End Sub

Private Sub tbEnterName_Change()
    cmdBigButton_Click 'bad idea, but someone will do it
End Sub

Private Sub UserForm_Click()
    cmdBigButton.Caption = ""Click This""
End Sub"
            };
            inputOutput.ControlNames.Add("cmdBtn1");
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test, Ignore("")]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameControlFromEventHandlerReference()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "cmdBtn1", newName: "cmdBigButton");
            var inputOutput = new RenameTestModuleDefinition("UserForm1", ComponentType.UserForm)
            {
                Input =
                    @"Private Sub cmdBtn1_Click()
End Sub

Private Sub tbEnterName_Change()
    cmdBtn1_Cl|ick 'bad idea, but someone will do it
End Sub

Private Sub UserForm_Click()
    cmdBtn1.Caption = ""Click This""
End Sub",
                Expected =
                    @"Private Sub cmdBigButton_Click()
End Sub

Private Sub tbEnterName_Change()
    cmdBigButton_Click 'bad idea, but someone will do it
End Sub

Private Sub UserForm_Click()
    cmdBigButton.Caption = ""Click This""
End Sub"
            };
            inputOutput.ControlNames.Add("cmdBtn1");
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test, Ignore("")]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameControlHandlesUnderscoresInNewName()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "bigButton_ClickAgain", newName: "bigButton_ClickAgain_AndAgain");
            var inputOutput = new RenameTestModuleDefinition("UserForm1", ComponentType.UserForm)
            {
                Input =
                    @"Private Sub bigBut|ton_ClickAgain_Click()
End Sub",
                Expected =
                    @"Private Sub bigButton_ClickAgain_AndAgain_Click()
End Sub"
            };
            inputOutput.ControlNames.Add("bigButton_ClickAgain");
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test, Ignore("")]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameControlSimilarNames()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "bigButton", newName: "smallButton");
            var inputOutput = new RenameTestModuleDefinition("UserForm1", ComponentType.UserForm)
            {
                Input =
                    @"Private Sub bigBu|tton_Click()
End Sub

Private Sub bigButton_Changed()
End Sub

Private Sub bigButton_Click_Click()
End Sub",
                Expected =
                    @"Private Sub smallButton_Click()
End Sub

Private Sub smallButton_Changed()
End Sub

Private Sub bigButton_Click_Click()
End Sub"
            };
            inputOutput.ControlNames.Add("bigButton");
            inputOutput.ControlNames.Add("bigButton_Click");
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenamePresenter_WarnsAboutControlEventHandlerRename_AbortsOnDeniedConfirmation()
        {
            var qmn = new QualifiedModuleName("TestProject", string.Empty, "TestComponent");
            var testDeclaration = new FunctionDeclaration(new QualifiedMemberName(qmn, "Foo"), null, null, "Variant", null, string.Empty, Accessibility.Public, null, null, Selection.Home, false, true, Enumerable.Empty<IParseTreeAnnotation>(), null);
            var originalTargetDeclaration = new FunctionDeclaration(new QualifiedMemberName(qmn, "Bar"), null, null, "Variant", null, string.Empty, Accessibility.Public, null, null, Selection.Home, false, true, Enumerable.Empty<IParseTreeAnnotation>(), null);

            var model = new RenameModel(originalTargetDeclaration)
            {
                Target = testDeclaration,
                IsControlEventHandlerRename = true,
                NewName = "FooBar"
            };

            Assert_WarnsAboutTargetMove_DenialOfConfirmationLeadsToAbort(model, confirm: false);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenamePresenter_WarnsAboutControlEventHandlerRename_ContinuesAfterConfirmation()
        {
            var qmn = new QualifiedModuleName("TestProject", string.Empty, "TestComponent");
            var testDeclaration = new FunctionDeclaration(new QualifiedMemberName(qmn, "Foo"), null, null, "Variant", null, string.Empty, Accessibility.Public, null, null, Selection.Home, false, true, Enumerable.Empty<IParseTreeAnnotation>(), null);
            var originalTargetDeclaration = new FunctionDeclaration(new QualifiedMemberName(qmn, "Bar"), null, null, "Variant", null, string.Empty, Accessibility.Public, null, null, Selection.Home, false, true, Enumerable.Empty<IParseTreeAnnotation>(), null);

            var model = new RenameModel(originalTargetDeclaration)
            {
                Target = testDeclaration,
                IsControlEventHandlerRename = true,
                NewName = "FooBar"
            };

            Assert_WarnsAboutTargetMove_DenialOfConfirmationLeadsToAbort(model, confirm: true);
        }

        #endregion
        #region Rename Event Tests

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameEvent()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "Foo", newName: "Goo");
            var inputOutput1 = new RenameTestModuleDefinition("Class1")
            {
                Input =
                    @"Public Event Fo|o(ByVal arg1 As Integer, ByVal arg2 As String)",

                Expected =
                    @"Public Event Goo(ByVal arg1 As Integer, ByVal arg2 As String)"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameEventAndHandlers()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "Foo", newName: "Goo");
            var inputOutput1 = new RenameTestModuleDefinition("Class1")
            {
                Input =
                    @"Public Event Fo|o(ByVal arg1 As Integer, ByVal arg2 As String)",

                Expected =
                    @"Public Event Goo(ByVal arg1 As Integer, ByVal arg2 As String)"
            };
            var inputOutput2 = new RenameTestModuleDefinition("Class2")
            {
                Input =
                    @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal i As Integer, ByVal s As String)
End Sub",
                Expected =
                    @"Private WithEvents abc As Class1

Private Sub abc_Goo(ByVal i As Integer, ByVal s As String)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutput2);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameEventUnaffectedByLookAlikeName()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "abc_Foo", newName: "abc_Goo");
            var inputOutput1 = new RenameTestModuleDefinition("Class1")
            {
                Input =
                    @"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)",

                Expected =
                    @"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)"
            };
            var inputOutput2 = new RenameTestModuleDefinition("Class2")
            {   //Note: no withEvents declaration, abc_Foo is just a Sub
                Input =
                    @"Private Sub abc_Fo|o(ByVal i As Integer, ByVal s As String)
End Sub",
                Expected =
                    @"Private Sub abc_Goo(ByVal i As Integer, ByVal s As String)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutput2);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameEventUnaffectedByLookAlikeName2()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "def_Foo", newName: "def_Goo");
            var inputOutput1 = new RenameTestModuleDefinition("Class1")
            {
                Input =
                    @"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)",

                Expected =
                    @"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)"
            };
            var inputOutput2 = new RenameTestModuleDefinition("Class2")
            {
                Input =
                    @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal i As Integer, ByVal s As String)
End Sub

Private Sub def_F|oo(ByVal i As Integer, ByVal s As String)
End Sub",
                Expected =
                    @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal i As Integer, ByVal s As String)
End Sub

Private Sub def_Goo(ByVal i As Integer, ByVal s As String)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutput2);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameEventAndHandlersNarrowScope()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "Foo", newName: "Goo");
            var inputOutputWithSelection = new RenameTestModuleDefinition("EventClass1")
            {
                Input =
                    @"Public Event Fo|o(ByVal arg1 As Integer, ByVal arg2 As String)
Public Event Bar()",

                Expected =
                    @"Public Event Goo(ByVal arg1 As Integer, ByVal arg2 As String)
Public Event Bar()"
            };
            var inputOutput2 = new RenameTestModuleDefinition("EventClass2")
            {
                Input =
                    @"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)",

                Expected =
                    @"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)",
            };
            var inputOutput3 = new RenameTestModuleDefinition("WithEvents1")
            {
                Input =
                    @"Private WithEvents abc As EventClass1
Private WithEvents otherEvents As EventClass2

Private Sub abc_Foo(ByVal i As Integer, ByVal s As String)
End Sub

Private Sub abc_Bar()
End Sub

Private Sub otherEvents_Foo(ByVal i As Integer, ByVal s As String)
End Sub",
                Expected =
                    @"Private WithEvents abc As EventClass1
Private WithEvents otherEvents As EventClass2

Private Sub abc_Goo(ByVal i As Integer, ByVal s As String)
End Sub

Private Sub abc_Bar()
End Sub

Private Sub otherEvents_Foo(ByVal i As Integer, ByVal s As String)
End Sub"
            };
            var inputOutput4 = new RenameTestModuleDefinition("WithEvents2")
            {
                Input =
                    @"Private WithEvents myEvents As EventClass1
Private WithEvents evenMoreEvents As EventClass2

Private Sub myEvents_Foo(ByVal i As Integer, ByVal s As String)
End Sub

Private Sub myEvents_Bar()
End Sub

Private Sub evenMoreEvents_Foo(ByVal i As Integer, ByVal s As String)
End Sub",
                Expected =
                    @"Private WithEvents myEvents As EventClass1
Private WithEvents evenMoreEvents As EventClass2

Private Sub myEvents_Goo(ByVal i As Integer, ByVal s As String)
End Sub

Private Sub myEvents_Bar()
End Sub

Private Sub evenMoreEvents_Foo(ByVal i As Integer, ByVal s As String)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutputWithSelection, inputOutput2, inputOutput3, inputOutput4);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameEventUpdatesUsages()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "MyEvent", newName: "YourEvent");
            var inputOutput1 = new RenameTestModuleDefinition("CEventClass")
            {
                Input =
                    @"
Public Event MyEv|ent(IDNumber As Long, ByRef Cancel As Boolean)

Sub AAA()
    RaiseEvent MyEvent(1234, False)
End Sub",
                Expected =
                    @"
Public Event YourEvent(IDNumber As Long, ByRef Cancel As Boolean)

Sub AAA()
    RaiseEvent YourEvent(1234, False)
End Sub"
            };
            var inputOutput2 = new RenameTestModuleDefinition("Class2")
            {
                Input =
                    @"
Private WithEvents XLEvents As CEventClass

Private Sub Class_Initialize()
    Set XLEvents = New CEventClass
End Sub

Private Sub XLEvents_MyEvent(IDNumber As Long, Cancel As Boolean)
    Cancel = True
End Sub",
                Expected =
                    @"
Private WithEvents XLEvents As CEventClass

Private Sub Class_Initialize()
    Set XLEvents = New CEventClass
End Sub

Private Sub XLEvents_YourEvent(IDNumber As Long, Cancel As Boolean)
    Cancel = True
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutput2);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameEventUsingWithEventsVariable()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "XLEvents", newName: "NewEventImpl");
            var inputOutput1 = new RenameTestModuleDefinition("CEventClass")
            {
                Input =
                    @"Public Event MyEvent(IDNumber As Long, ByRef Cancel As Boolean)

Sub AAA()
    RaiseEvent MyEvent(1234, False)
End Sub",
                Expected =
                    @"Public Event MyEvent(IDNumber As Long, ByRef Cancel As Boolean)

Sub AAA()
    RaiseEvent MyEvent(1234, False)
End Sub"
            };

            var inputOutputWithRenameTarget = new RenameTestModuleDefinition("Class2")
            {
                Input =
                    @"Private WithEvents XLEve|nts As CEventClass

Private Sub Class_Initialize()
    Set XLEvents = New CEventClass
End Sub

Private Sub XLEvents_MyEvent(IDNumber As Long, Cancel As Boolean)
    Cancel = True
End Sub",
                Expected =
                    @"Private WithEvents NewEventImpl As CEventClass

Private Sub Class_Initialize()
    Set NewEventImpl = New CEventClass
End Sub

Private Sub NewEventImpl_MyEvent(IDNumber As Long, Cancel As Boolean)
    Cancel = True
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutputWithRenameTarget);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameEventUsingWithEventsVariableConfictingName()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "abc", newName: "def");
            var inputOutput1 = new RenameTestModuleDefinition("Class1")
            {
                Input =
                    @"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)",

                Expected =
                    @"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)"
            };
            var inputOutput2 = new RenameTestModuleDefinition("Class2")
            {
                Input =
                    @"Private WithEvents a|bc As Class1

Private Sub abc_Foo(ByVal i As Integer, ByVal s As String)
End Sub

Private Sub abc_HorriblyNamedSub()
End Sub",
                Expected =
                    @"Private WithEvents def As Class1

Private Sub def_Foo(ByVal i As Integer, ByVal s As String)
End Sub

Private Sub abc_HorriblyNamedSub()
End Sub",
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutput2);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameEventFromHandler()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "MyEvent", newName: "YourEvent_withUnderscore");
            var inputOutput1 = new RenameTestModuleDefinition("CEventClass")
            {
                Input =
                    @"
Public Event MyEvent(IDNumber As Long, ByRef Cancel As Boolean)

Sub AAA()
    RaiseEvent MyEvent(1234, False)
End Sub",
                Expected =
                    @"
Public Event YourEvent_withUnderscore(IDNumber As Long, ByRef Cancel As Boolean)

Sub AAA()
    RaiseEvent YourEvent_withUnderscore(1234, False)
End Sub"
            };

            var inputOutput2 = new RenameTestModuleDefinition("Class2")
            {
                Input =
                    @"Private WithEvents XLEvents As CEventClass

Private Sub Class_Initialize()
    Set XLEvents = New CEventClass
End Sub

Private Sub XLEvents_My|Event(IDNumber As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Function DumbFunction() As Long
    XLEvents_MyEvent 6,wasCancelled
    DumbFunction = 8
End Function",

                Expected =
                    @"Private WithEvents XLEvents As CEventClass

Private Sub Class_Initialize()
    Set XLEvents = New CEventClass
End Sub

Private Sub XLEvents_YourEvent_withUnderscore(IDNumber As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Function DumbFunction() As Long
    XLEvents_YourEvent_withUnderscore 6,wasCancelled
    DumbFunction = 8
End Function"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutput2);

            Assert.IsTrue(tdo.Model.IsUserEventHandlerRename);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenamePresenter_WarnsAboutEventHandlerRename_AbortsOnDeniedConfirmation()
        {
            var qmn = new QualifiedModuleName("TestProject", string.Empty, "TestComponent");
            var testDeclaration = new FunctionDeclaration(new QualifiedMemberName(qmn, "Foo"), null, null, "Variant", null, string.Empty, Accessibility.Public, null, null, Selection.Home, false, true, Enumerable.Empty<IParseTreeAnnotation>(), null);
            var originalTargetDeclaration = new FunctionDeclaration(new QualifiedMemberName(qmn, "Bar"), null, null, "Variant", null, string.Empty, Accessibility.Public, null, null, Selection.Home, false, true, Enumerable.Empty<IParseTreeAnnotation>(), null);

            var model = new RenameModel(originalTargetDeclaration)
            {
                Target = testDeclaration,
                IsUserEventHandlerRename = true,
                NewName = "FooBar"
            };

            Assert_WarnsAboutTargetMove_DenialOfConfirmationLeadsToAbort(model, confirm: false);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenamePresenter_WarnsAboutEventHandlerRename_ContinuesAfterConfirmation()
        {
            var qmn = new QualifiedModuleName("TestProject", string.Empty, "TestComponent");
            var testDeclaration = new FunctionDeclaration(new QualifiedMemberName(qmn, "Foo"), null, null, "Variant", null, string.Empty, Accessibility.Public, null, null, Selection.Home, false, true, Enumerable.Empty<IParseTreeAnnotation>(), null);
            var originalTargetDeclaration = new FunctionDeclaration(new QualifiedMemberName(qmn, "Bar"), null, null, "Variant", null, string.Empty, Accessibility.Public, null, null, Selection.Home, false, true, Enumerable.Empty<IParseTreeAnnotation>(), null);

            var model = new RenameModel(originalTargetDeclaration)
            {
                Target = testDeclaration,
                IsUserEventHandlerRename = true,
                NewName = "FooBar"
            };

            Assert_WarnsAboutTargetMove_DenialOfConfirmationLeadsToAbort(model, confirm: true);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameEventFromUsage()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "MyEvent", newName: "YourEvent");
            var inputOutput1 = new RenameTestModuleDefinition("CEventClass")
            {
                Input =
                    @"
Public Event MyEvent(IDNumber As Long, ByRef Cancel As Boolean)

Sub AAA()
    RaiseEvent My|Event(1234, False)
End Sub",
                Expected =
                    @"
Public Event YourEvent(IDNumber As Long, ByRef Cancel As Boolean)

Sub AAA()
    RaiseEvent YourEvent(1234, False)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1);
        }

        #endregion
        #region Rename Interface Tests

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameInterface()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "DoSomething", newName: "DoNothing");
            var inputOutput1 = new RenameTestModuleDefinition("IClass1")
            {
                Input =
                    @"Public Sub DoSo|mething(ByVal a As Integer, ByVal b As String)
End Sub",
                Expected =
                    @"Public Sub DoNothing(ByVal a As Integer, ByVal b As String)
End Sub"
            };
            var inputOutput2 = new RenameTestModuleDefinition("Class1")
            {
                Input =
                    @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub",
                Expected =
                    @"Implements IClass1

Private Sub IClass1_DoNothing(ByVal a As Integer, ByVal b As String)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutput2);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameInterfaceMemberDuplicateMemberInOtherInterface()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "DoSomething", newName: "DoNothing");
            var inputOutput1 = new RenameTestModuleDefinition("IClass1")
            {
                Input =
                    @"Public Sub DoS|omething()
End Sub",
                Expected =
                    @"Public Sub DoNothing()
End Sub"
            };
            var inputOutput2 = new RenameTestModuleDefinition("IClass2")
            {
                Input =
                    @"Public Sub DoSomething()
End Sub",
                Expected =
                    @"Public Sub DoSomething()
End Sub"
            };
            var inputOutput3 = new RenameTestModuleDefinition("Class1")
            {
                Input =
                    @"Implements IClass1

Private Sub IClass1_DoSomething()
End Sub",
                CheckExpectedEqualsActual = false
            };
            var inputOutput4 = new RenameTestModuleDefinition("Class2")
            {
                Input =
                    @"Implements IClass2

Private Sub IClass2_DoSomething()
End Sub",
                CheckExpectedEqualsActual = false
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutput2, inputOutput3, inputOutput4);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameInterfaceReferences()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "DoSomething", newName: "DoNothing");
            var inputOutputWithSelection = new RenameTestModuleDefinition("IClass1")
            {
                Input =
                    @"Public Sub DoS|omething()
End Sub",
                Expected =
                    @"Public Sub DoNothing()
End Sub"
            };
            var inputOutput2 = new RenameTestModuleDefinition("Class1")
            {
                Input =
                    @"Implements IClass1

Private Sub IClass1_DoSomething()
End Sub",
                Expected =
                    @"Implements IClass1

Private Sub IClass1_DoNothing()
End Sub"
            };
            var inputOutput3 = new RenameTestModuleDefinition("Class2")
            {
                Input =
                    @"Private Sub RefTheInterface()
    Dim c1 As Class1
    Set c1 = new IClass1
    c1.DoSomething
End Sub

Private Sub RefTheInterface2()
    Dim c1 As Class1
    Dim c2 As IClass1
    Set c1 = new Class1
    Set c2 = c1
    c1.DoSomething
End Sub",
                Expected =
                    @"Private Sub RefTheInterface()
    Dim c1 As Class1
    Set c1 = new IClass1
    c1.DoNothing
End Sub

Private Sub RefTheInterface2()
    Dim c1 As Class1
    Dim c2 As IClass1
    Set c1 = new Class1
    Set c2 = c1
    c1.DoNothing
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutputWithSelection, inputOutput2, inputOutput3);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameInterfaceFromImplementingMember()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "DoSomething", newName: "DoNothing");
            var inputOutput1 = new RenameTestModuleDefinition("IClass1")
            {
                Input =
                    @"Public Sub DoSomething()
End Sub",
                Expected =
                    @"Public Sub DoNothing()
End Sub"
            };
            var inputOutputWithSelection = new RenameTestModuleDefinition("Class1")
            {
                Input =
                    @"Implements IClass1

Private Sub IC|lass1_DoSomething()
End Sub",
                Expected =
                    @"Implements IClass1

Private Sub IClass1_DoNothing()
End Sub"
            };
            var inputOutput3 = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Private Sub RefTheInterface()
    Dim c1 As Class1
    Set c1 = new IClass1
    c1.DoSomething
End Sub

Private Sub RefTheInterface2()
    Dim c1 As Class1
    Dim c2 As IClass1
    Set c1 = new Class1
    Set c2 = c1
    c1.DoSomething
End Sub",
                Expected =
                    @"Private Sub RefTheInterface()
    Dim c1 As Class1
    Set c1 = new IClass1
    c1.DoNothing
End Sub

Private Sub RefTheInterface2()
    Dim c1 As Class1
    Dim c2 As IClass1
    Set c1 = new Class1
    Set c2 = c1
    c1.DoNothing
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutputWithSelection, inputOutput3);

            Assert.IsTrue(tdo.Model.IsInterfaceMemberRename);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenamePresenter_WarnsAboutInterfaceVariableImplementationRename_AbortsOnDeniedConfirmation()
        {
            var qmn = new QualifiedModuleName("TestProject", string.Empty, "TestComponent");
            var testDeclaration = new FunctionDeclaration(new QualifiedMemberName(qmn,"Foo"), null, null, "Variant", null, string.Empty, Accessibility.Public, null, null, Selection.Home, false, true, Enumerable.Empty<IParseTreeAnnotation>(), null);
            var originalTargetDeclaration = new FunctionDeclaration(new QualifiedMemberName(qmn, "Bar"), null, null, "Variant", null, string.Empty, Accessibility.Public, null, null, Selection.Home, false, true, Enumerable.Empty<IParseTreeAnnotation>(), null);

            var model = new RenameModel(originalTargetDeclaration)
            {
                Target = testDeclaration,
                IsInterfaceMemberRename = true,
                NewName = "FooBar"
            };

            Assert_WarnsAboutTargetMove_DenialOfConfirmationLeadsToAbort(model, confirm: false);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenamePresenter_WarnsAboutInterfaceVariableImplementationRename_ContinuesAfterConfirmation()
        {
            var qmn = new QualifiedModuleName("TestProject", string.Empty, "TestComponent");
            var testDeclaration = new FunctionDeclaration(new QualifiedMemberName(qmn, "Foo"), null, null, "Variant", null, string.Empty, Accessibility.Public, null, null, Selection.Home, false, true, Enumerable.Empty<IParseTreeAnnotation>(), null);
            var originalTargetDeclaration = new FunctionDeclaration(new QualifiedMemberName(qmn, "Bar"), null, null, "Variant", null, string.Empty, Accessibility.Public, null, null, Selection.Home, false, true, Enumerable.Empty<IParseTreeAnnotation>(), null);

            var model = new RenameModel(originalTargetDeclaration)
            {
                Target = testDeclaration,
                IsInterfaceMemberRename = true,
                NewName = "FooBar"
            };

            Assert_WarnsAboutTargetMove_DenialOfConfirmationLeadsToAbort(model, confirm: true);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameInterfaceVariable()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "Foo", newName: "Bar");
            var inputOutput1 = new RenameTestModuleDefinition("IClass1")
            {
                Input =
                    @"Public F|oo As Long",
                Expected =
                    @"Public Bar As Long"
            };
            var inputOutput2 = new RenameTestModuleDefinition("Class1")
            {
                Input =
                    @"Implements IClass1

Private Property Get IClass1_Foo() As Long
End Property

Private Property Let IClass1_Foo(rhs As Long)
End Property",
                Expected =
                    @"Implements IClass1

Private Property Get IClass1_Bar() As Long
End Property

Private Property Let IClass1_Bar(rhs As Long)
End Property"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutput2);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameInterfaceVariable_AcceptPrompt()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "Foo", newName: "Bar");
            var inputOutput1 = new RenameTestModuleDefinition("Class1")
            {
                Input = @"Implements IClass1

Private Property Get IClass1_F|oo() As Long
End Property

Private Property Let IClass1_Foo(rhs As Long)
End Property",
                Expected =
                    @"Implements IClass1

Private Property Get IClass1_Bar() As Long
End Property

Private Property Let IClass1_Bar(rhs As Long)
End Property"
            };

            var inputOutput2 = new RenameTestModuleDefinition("IClass1")
            {
                Input =
                    @"Public Foo As Long",
                Expected =
                    @"Public Bar As Long"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutput2);

            Assert.IsTrue(tdo.Model.IsInterfaceMemberRename);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameInterfaceFromMemberProperty()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "Something", newName: "Nothing");
            var inputOutput1 = new RenameTestModuleDefinition("IClass1")
            {
                Input =
                    @"Public Property Set Something(arg1 As Long)
End Property

Public Property Get Something() As Long
End Property",
                Expected =
                    @"Public Property Set Nothing(arg1 As Long)
End Property

Public Property Get Nothing() As Long
End Property"
            };

            var inputOutputWithSelection = new RenameTestModuleDefinition("Class1")
            {
                Input =
                    @"Implements IClass1

Private Property Set IClass1_Some|thing(arg1 As Long)
End Property

Private Property Get IClass1_Something() As Long
End Property",
                Expected =
                    @"Implements IClass1

Private Property Set IClass1_Nothing(arg1 As Long)
End Property

Private Property Get IClass1_Nothing() As Long
End Property"
            };

            var inputOutput3 = new RenameTestModuleDefinition("Class2")
            {
                Input =
                    @"Private Sub RefTheInterface()
    Dim c1 As Class1
    Set c1 = new IClass1
    c1.Something 7
End Sub

Private Sub RefTheInterface2()
    Dim c1 As Class1
    Dim c2 As IClass1
    Set c1 = new Class1
    Set c2 = c1
    c1.Something 7
End Sub",
                Expected =
                    @"Private Sub RefTheInterface()
    Dim c1 As Class1
    Set c1 = new IClass1
    c1.Nothing 7
End Sub

Private Sub RefTheInterface2()
    Dim c1 As Class1
    Dim c2 As IClass1
    Set c1 = new Class1
    Set c2 = c1
    c1.Nothing 7
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutputWithSelection, inputOutput3);

            Assert.IsTrue(tdo.Model.IsInterfaceMemberRename);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameInterfaceNoImplementers()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "DoSomething", newName: "DoNothing");
            var inputOutput1 = new RenameTestModuleDefinition("IClass1")
            {
                Input =
                    @"Public Sub Do|Something()
End Sub",
                Expected =
                    @"Public Sub DoNothing()
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameInterfaceFromReference()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "DoSomething", newName: "DoNothing");
            var inputOutput1 = new RenameTestModuleDefinition("IClass1")
            {
                Input =
                    @"Public Sub DoSomething(arg1 As Long)
End Sub",
                Expected =
                    @"Public Sub DoNothing(arg1 As Long)
End Sub",
            };

            var inputOutput2 = new RenameTestModuleDefinition("Class1")
            {
                Input =
                    @"Implements IClass1

Private Sub IClass1_DoSomething(arg1 As Long)
End Sub",
                Expected =
                    @"Implements IClass1

Private Sub IClass1_DoNothing(arg1 As Long)
End Sub"
            };

            var inputOutputWithSelection = new RenameTestModuleDefinition("Class2")
            {
                Input =
                    @"Private Sub RefTheInterface()
    Dim c1 As Class1
    Set c1 = new IClass1
    c1.DoS|omething
End Sub

Private Sub RefTheInterface2()
    Dim c3 As Class1
    Dim c2 As IClass1
    Set c3 = new Class1
    Set c2 = c3
    c3.DoSomething
End Sub",
                Expected =
                    @"Private Sub RefTheInterface()
    Dim c1 As Class1
    Set c1 = new IClass1
    c1.DoNothing
End Sub

Private Sub RefTheInterface2()
    Dim c3 As Class1
    Dim c2 As IClass1
    Set c3 = new Class1
    Set c2 = c3
    c3.DoNothing
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutput2, inputOutputWithSelection);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameInterfaceReferencesWithinScope()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "DoSomething", newName: "DoNothing");
            var inputOutputWithSelection = new RenameTestModuleDefinition("IClass1")
            {
                Input =
                    @"Public Sub DoSo|mething()
End Sub",
                Expected =
                    @"Public Sub DoNothing()
End Sub"
            };

            var inputOutput2 = new RenameTestModuleDefinition("Class1")
            {
                Input =
                    @"Implements IClass1

Private Sub IClass1_DoSomething()
End Sub",
                Expected =
                    @"Implements IClass1

Private Sub IClass1_DoNothing()
End Sub"
            };

            var inputOutput3 = new RenameTestModuleDefinition("Class2")
            {
                Input =
                    @"Private Sub RefTheInterface()
    Dim c1 As Class1
    Set c1 = new IClass1
    c1.DoSomething
End Sub

Private Sub RefTheInterface2()
    Dim c1 As Class2
    Dim c2 As IClass1
    Set c1 = new Class2
    Set c2 = c1
    c1.DoSomething  'This is left alone because it is a member of Class2, not the interface
    c2.DoSomething
End Sub",
                Expected =
                    @"Private Sub RefTheInterface()
    Dim c1 As Class1
    Set c1 = new IClass1
    c1.DoNothing
End Sub

Private Sub RefTheInterface2()
    Dim c1 As Class2
    Dim c2 As IClass1
    Set c1 = new Class2
    Set c2 = c1
    c1.DoSomething  'This is left alone because it is a member of Class2, not the interface
    c2.DoNothing
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutputWithSelection, inputOutput2, inputOutput3);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameInterface_AcceptPrompt()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "DoSomething", newName: "DoNothing");
            var inputOutput1 = new RenameTestModuleDefinition("Class1")
            {
                Input =
                    @"Implements IClass1

Private Sub ICla|ss1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub",
                Expected =
                    @"Implements IClass1

Private Sub IClass1_DoNothing(ByVal a As Integer, ByVal b As String)
End Sub"
            };

            var inputOutput2 = new RenameTestModuleDefinition("IClass1")
            {
                Input =
                    @"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub",
                Expected =
                    @"Public Sub DoNothing(ByVal a As Integer, ByVal b As String)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutput2);

            Assert.IsTrue(tdo.Model.IsInterfaceMemberRename);
        }

        #endregion
        #region Rename CodeModule Tests

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameModuleFromImplementsStmt()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "IClass1", newName: "INewClass");
            var inputOutput1 = new RenameTestModuleDefinition("IClass1")
            {
                Input =
                    @"Public Sub DoSomething()
End Sub",
                CheckExpectedEqualsActual = false
            };
            var inputOutputWithSelection = new RenameTestModuleDefinition("Class1")
            {
                Input =
                    @"Implements ICl|ass1

Private Sub IClass1_DoSomething()
End Sub",
                Expected =
                    @"Implements INewClass

Private Sub INewClass_DoSomething()
End Sub"
            };
            var inputOutput3 = new RenameTestModuleDefinition("Class2")
            {
                Input =
                    @"Implements IClass1

Private Sub IClass1_DoSomething()
End Sub",
                Expected =
                    @"Implements INewClass

Private Sub INewClass_DoSomething()
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutputWithSelection, inputOutput3);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameModuleFromReference()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "CTestClass", newName: "CMyTestClass");
            var inputOutput1 = new RenameTestModuleDefinition("CTestClass")
            {
                Input =
                    @"
Sub Foo()
End Sub",
                NewName = "CMyTestClass"
            };
            inputOutput1.Expected = inputOutput1.Input;

            var inputOutput2 = new RenameTestModuleDefinition("Class2")
            {
                Input =

                    @"
Sub Foo2()
    Dim c1 As CTes|tClass
    Set c1 = new CTestClass
    c1.Foo
End Sub",
                Expected =
                    @"
Sub Foo2()
    Dim c1 As CMyTestClass
    Set c1 = new CMyTestClass
    c1.Foo
End Sub"
            };

            //This will fail with because of an invalid dictionary access if the class rename does not succeed.
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutput2);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameCodeModule()
        {
            const string newName = "RenameModule";

            const string inputCode =
                @"Private Sub Foo(ByVal a As Integer, ByVal b As String)
End Sub";

            var tdo = new RenameTestsDataObject("Class1", DeclarationType.ClassModule, newName);
            var testModuleDefinition = new RenameTestModuleDefinition("Class1")
            {
                Input = inputCode,
                NewName = newName
            };

            //This will fail with because of an invalid dictionary access if the class rename does not succeed.
            PerformExpectedVersusActualRenameTests(tdo, testModuleDefinition);
        }

        #endregion
        #region Rename Project Tests

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameProject()
        {
            const string newName = "Renamed";

            var presenterAction = AdjustName(newName);

            var vbe = TestVbe(string.Empty , out _);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var target = state.DeclarationFinder.UserDeclarations(DeclarationType.Project).Single();

                var refactoring = TestRefactoring(rewritingManager, state, presenterAction);

                refactoring.Refactor(target);

                Assert.AreEqual(newName, vbe.VBProjects[0].Name);
            }
        }

        #endregion
        #region Rename Enumeration Tests

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameEnumeration()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "FruitType", newName: "Fruits");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Option Explicit

Public Enum Frui|tType
    Apple = 1
    Orange = 2
    Plum = 3
End Enum

Sub DoSomething()
    MsgBox CStr(FruitType.Apple)
End Sub",
                Expected =
                    @"Option Explicit

Public Enum Fruits
    Apple = 1
    Orange = 2
    Plum = 3
End Enum

Sub DoSomething()
    MsgBox CStr(Fruits.Apple)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameEnumerationMember()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "Apple", newName: "CranApple");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Option Explicit

Public Enum FruitType
    App|le = 1
    Orange = 2
    Plum = 3
End Enum

Sub DoSomething()
    MsgBox CStr(Apple)
End Sub",
                Expected =
                    @"Option Explicit

Public Enum FruitType
    CranApple = 1
    Orange = 2
    Plum = 3
End Enum

Sub DoSomething()
    MsgBox CStr(CranApple)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameEnumerationMember_WarnMemberExists_CancelIfNotAccepted()
        {
            var moduleCode =
                @"Option Explicit

Public Enum FruitType
    App|le = 1
    Orange = 2
    Plum = 3
End Enum

Sub DoSomething()
    MsgBox CStr(Apple)
End Sub";

            Assert_WarnsAboutRenameConflict_ConfirmationOutcomeHasExpectedResult(
                moduleCode,
                ComponentType.StandardModule,
                "Apple",
                DeclarationType.EnumerationMember,
                "Plum",
                false,
                RefactoringDialogResult.Cancel);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameEnumerationMember_WarnMemberExists_ProceedIfAccepted()
        {
            var moduleCode =
                @"Option Explicit

Public Enum FruitType
    App|le = 1
    Orange = 2
    Plum = 3
End Enum

Sub DoSomething()
    MsgBox CStr(Apple)
End Sub";

            Assert_WarnsAboutRenameConflict_ConfirmationOutcomeHasExpectedResult(
                moduleCode,
                ComponentType.StandardModule,
                "Apple",
                DeclarationType.EnumerationMember,
                "Plum",
                true,
                RefactoringDialogResult.Execute);
        }

        #endregion
        #region Rename UDT Tests

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenamePublicUDT()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "UserType", newName: "NewUserType");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Option Explicit

Public Type UserType|
    foo As String
    bar As Long
End Type


Private Sub DoSomething(baz As UserType)
    MsgBox CStr(baz.bar)
End Sub",
                Expected =
                    @"Option Explicit

Public Type NewUserType
    foo As String
    bar As Long
End Type


Private Sub DoSomething(baz As NewUserType)
    MsgBox CStr(baz.bar)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenamePrivateUDT()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "UserType", newName: "NewUserType");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Option Explicit

Public Type UserType|
    foo As String
    bar As Long
End Type


Private Sub DoSomething(baz As UserType)
    MsgBox CStr(baz.bar)
End Sub",
                Expected =
                    @"Option Explicit

Public Type NewUserType
    foo As String
    bar As Long
End Type


Private Sub DoSomething(baz As NewUserType)
    MsgBox CStr(baz.bar)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameUDTMember()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "bar", newName: "fooBar");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Option Explicit

Private Type UserType
    foo As String
    bar| As Long
End Type


Private Sub DoSomething(baz As UserType)
    MsgBox CStr(baz.bar)
End Sub",
                Expected =
                    @"Option Explicit

Private Type UserType
    foo As String
    fooBar As Long
End Type


Private Sub DoSomething(baz As UserType)
    MsgBox CStr(baz.fooBar)
End Sub"
            };

            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameUDTMember_WarnMemberExists_CancelIfNotAccepted()
        {
            var moduleCode =
@"Option Explicit

Private Type UserType
    foo As String
    bar| As Long
End Type


Private Sub DoSomething(baz As UserType)
    MsgBox CStr(baz.bar)
End Sub";

            Assert_WarnsAboutRenameConflict_ConfirmationOutcomeHasExpectedResult(
                moduleCode,
                ComponentType.StandardModule,
                "bar",
                DeclarationType.UserDefinedTypeMember,
                "foo",
                false,
                RefactoringDialogResult.Cancel);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameUDTMember_WarnMemberExists_ProceedIfAccepted()
        {
            var moduleCode =
                @"Option Explicit

Private Type UserType
    foo As String
    bar| As Long
End Type


Private Sub DoSomething(baz As UserType)
    MsgBox CStr(baz.bar)
End Sub";
            Assert_WarnsAboutRenameConflict_ConfirmationOutcomeHasExpectedResult(
                moduleCode,
                ComponentType.StandardModule,
                "bar",
                DeclarationType.UserDefinedTypeMember,
                "foo",
                true,
                RefactoringDialogResult.Execute);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenamePublicUDT_ReferenceInDifferentModule()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "UserType", newName: "NewUserType");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Option Explicit

Public Type UserType|
    foo As String
    bar As Long
End Type",

                Expected =
                    @"Option Explicit

Public Type NewUserType
    foo As String
    bar As Long
End Type"
            };

            var otherModule = new RenameTestModuleDefinition("Module2", ComponentType.StandardModule)
            {
                Input =
                    @"Option Explicit

Private Sub DoSomething(baz As UserType)
    MsgBox CStr(baz.bar)
End Sub",
                Expected =
                    @"Option Explicit

Private Sub DoSomething(baz As NewUserType)
    MsgBox CStr(baz.bar)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput, otherModule);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenamePublicUDTMember_ReferenceInDifferentModule()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "bar", newName: "fooBar");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Option Explicit

Public Type UserType
    foo As String
    bar| As Long
End Type",
                Expected =
                    @"Option Explicit

Public Type UserType
    foo As String
    fooBar As Long
End Type"
            };

            var otherModule = new RenameTestModuleDefinition("Module2", ComponentType.StandardModule)
            {
                Input =
                    @"Option Explicit

Private Sub DoSomething(baz As UserType)
    MsgBox CStr(baz.bar)
End Sub",
                Expected =
                    @"Option Explicit

Private Sub DoSomething(baz As UserType)
    MsgBox CStr(baz.fooBar)
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput, otherModule);
        }

        #endregion
        #region Rename Label Tests
        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameLabel()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "EH", newName: "ErrorHandler");
            var inputOutput1 = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Option Explicit

Sub DoSomething()
    On Error goto EH
    Dim check As Double
    check = 1/0
    Exit Sub
E|H:
    MsgBox ""We had an error""
End Sub",
                Expected =
                    @"Option Explicit

Sub DoSomething()
    On Error goto ErrorHandler
    Dim check As Double
    check = 1/0
    Exit Sub
ErrorHandler:
    MsgBox ""We had an error""
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1);
        }
        #endregion
        #region Property Tests
        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RefactorProperties_UpdatesReferences()
        {
            const string oldName = "Column";
            const string refactoredName = "Rank";

            var classInputOutput = new RenameTestModuleDefinition("MyClass", ComponentType.ClassModule)
            {
                Input = $@"Option Explicit

Private colValue As Long

Public Property Get {oldName}() As Long
    {oldName} = colValue
End Property
Public Property Let {FAUX_CURSOR}{oldName}(value As Long)
    colValue = value
End Property
",
                Expected = $@"Option Explicit

Private colValue As Long

Public Property Get {refactoredName}() As Long
    {refactoredName} = colValue
End Property
Public Property Let {refactoredName}(value As Long)
    colValue = value
End Property
"
            };
            var usageInputOutput = new RenameTestModuleDefinition("Usage", ComponentType.StandardModule)
            {
                Input = $@"Option Explicit

Public Sub useColValue()
    Dim instance As MyClass
    Set instance = New MyClass
    instance.{oldName} = 97521
    PrintValue instance.{oldName} & ""is the value""
End Sub

Private Sub PrintValue(value As String)
    Debug.Print value
End Sub
",
                Expected = $@"Option Explicit

Public Sub useColValue()
    Dim instance As MyClass
    Set instance = New MyClass
    instance.{refactoredName} = 97521
    PrintValue instance.{refactoredName} & ""is the value""
End Sub

Private Sub PrintValue(value As String)
    Debug.Print value
End Sub
"
            };

            var tdo = new RenameTestsDataObject(oldName, DeclarationType.PropertyLet, refactoredName)
            {
                UseLibraries = true
            };

            PerformExpectedVersusActualRenameTests(tdo, classInputOutput, usageInputOutput);
        }

        //Issue: https://github.com/rubberduck-vba/Rubberduck/issues/4349
        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_DoesNotWarnForUDTMember()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "VS", newName: "VerySatisfiedResponses");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
@"Private Type TMonthScoreInfo
            VerySatisfiedResponses As Long
        End Type

        Private monthScoreInfo As TMonthScoreInfo

        Public Property Get V|S() As Long
            VS = monthScoreInfo.VerySatisfiedResponses
        End Property
        Public Property Let VS(ByVal theVal As Long)
            monthScoreInfo.VerySatisfiedResponses = theVal
        End Property",
                Expected =
@"Private Type TMonthScoreInfo
            VerySatisfiedResponses As Long
        End Type

        Private monthScoreInfo As TMonthScoreInfo

        Public Property Get VerySatisfiedResponses() As Long
            VerySatisfiedResponses = monthScoreInfo.VerySatisfiedResponses
        End Property
        Public Property Let VerySatisfiedResponses(ByVal theVal As Long)
            monthScoreInfo.VerySatisfiedResponses = theVal
        End Property"
            };

            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        //Issue: https://github.com/rubberduck-vba/Rubberduck/issues/4349
        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_DoesNotWarnForEnumMember()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "VerySatisfiedID", newName: "VerySatisfiedResponse");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
@"Private Enum MonthScoreTypes
            VerySatisfiedResponse
            VeryDissatisfiedResponse
        End Enum

        Public Property Get V|erySatisfiedID() As Long
            VerySatisfiedID = MonthScoreTypes.VerySatisfiedResponse
        End Property",
                Expected =
@"Private Enum MonthScoreTypes
            VerySatisfiedResponse
            VeryDissatisfiedResponse
        End Enum

        Public Property Get VerySatisfiedResponse() As Long
            VerySatisfiedResponse = MonthScoreTypes.VerySatisfiedResponse
        End Property",
            };

            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        //Issue: https://github.com/rubberduck-vba/Rubberduck/issues/4349
        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_DoesNotWarnForMember()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "VerySatisfiedResponse", newName: "VerySatisfiedID");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
@"Private Enum MonthScoreTypes
            VerySa|tisfiedResponse
            VeryDissatisfiedResponse
        End Enum

        Public Property Get VerySatisfiedID() As Long
            VerySatisfiedID = MonthScoreTypes.VerySatisfiedResponse
        End Property",
                Expected =
@"Private Enum MonthScoreTypes
            VerySatisfiedID
            VeryDissatisfiedResponse
        End Enum

        Public Property Get VerySatisfiedID() As Long
            VerySatisfiedID = MonthScoreTypes.VerySatisfiedID
        End Property",
            };

            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        #endregion
        #region Other Tests

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_CheckAllRefactorCallPaths()
        {
            RefactorParams[] refactorParams = { RefactorParams.None, RefactorParams.QualifiedSelection, RefactorParams.Declaration };
            foreach (var param in refactorParams)
            {
                var tdo = param == RefactorParams.Declaration
                    ? new RenameTestsDataObject("Foo", DeclarationType.Procedure, "Goo")
                    : new RenameTestsDataObject(selectedIdentifier: "Foo", newName: "Goo");

                var inputOutput = new RenameTestModuleDefinition("Class1")
                {
                    Input =
                        @"Private Sub F|oo()
End Sub",
                    Expected =
                        @"Private Sub Goo()
End Sub"
                };
                tdo.RefactorParamType = param;

                PerformExpectedVersusActualRenameTests(tdo, inputOutput);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void Rename_PresenterIsNull()
        {
            const string inputCode =
                @"Private Sub Foo()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var qualifiedSelection = new QualifiedSelection(component.QualifiedModuleName, Selection.Home);
                var factory = new Mock<IRefactoringPresenterFactory>();
                factory.Setup(m => m.Create<IRenamePresenter, RenameModel>(It.IsAny<RenameModel>())).Returns((RenameModel model) => null);
                var selectionService = MockedSelectionService();
                var refactoring = TestRefactoring(rewritingManager, state, factory.Object, selectionService);

                Assert.Throws<InvalidRefactoringPresenterException>(() => refactoring.Refactor(qualifiedSelection));

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(inputCode, actualCode);
            }
        }
      
        [Category("Refactorings")]
        [Category("Rename")]
        [TestCase("Class_Initialize")]
        [TestCase("Class_Terminate")]
        public void Rename_StandardEventHandler(string handlerName)
        {
            var inputCode =
                $@"Private Sub {handlerName}()
End Sub";

            var presenterAction = AdjustName("test");

            var actualCode = RefactoredCode(
                handlerName,
                DeclarationType.Procedure,
                presenterAction,
                typeof(TargetDeclarationIsStandardEventHandlerException),
                ("TestClass", inputCode, ComponentType.ClassModule));
            Assert.AreEqual(inputCode, actualCode["TestClass"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void Rename_NotUserDefined()
        {
            const string inputCode = 
                @"Private Sub Foo()
    Dim bar As Color|ScaleCriteria
End Sub";

            var tdo = new RenameTestsDataObject(declarationName: "ColorScaleCriteria", newName: "Goo", declarationType: DeclarationType.ClassModule);
            var inputOutput1 = new RenameTestModuleDefinition("Class1")
            {
                Input = inputCode,
                Expected = inputCode.Replace("|", string.Empty)
            };
            tdo.UseLibraries = true;
            tdo.ExpectedException = typeof(TargetDeclarationNotUserDefinedException);

            PerformExpectedVersusActualRenameTests(tdo, inputOutput1);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        [Ignore("Something is off with the project id of the implemented class: it does not agree with the project id of the exposing library.")]
        public void Rename_ImplementedInterfaceNotUserDefined()
        {
            var inputCode =
                @"Implements PivotFields

Private Property Get Pivo|tFields_Count() As Long
End Property

Private Function PivotFields_Item(Index) As Object
End Function

Private Property Get PivotFields_Application() As Application
End Property

Private Property Get PivotFields_Creator() As XlCreator
End Property

Private Property Get PivotFields_Parent() As PivotTable
End Property
";

            var tdo = new RenameTestsDataObject(declarationName: "PivotFields_Count", newName: "Goo", declarationType: DeclarationType.PropertyGet);
            var inputOutput1 = new RenameTestModuleDefinition("Class1")
            {
                Input = inputCode,
                Expected = inputCode.Replace("|", string.Empty)
            };
            tdo.UseLibraries = true;
            tdo.ExpectedException = typeof(TargetDeclarationNotUserDefinedException);

            PerformExpectedVersusActualRenameTests(tdo, inputOutput1);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void Rename_ModelIsNull()
        {
            const string inputCode =
                @"Private Sub Foo()
End Sub";

            Func<RenameModel, RenameModel> presenterAction = model => null;

            var actualCode = RefactoredCode(
                inputCode, 
                "Foo", 
                DeclarationType.Procedure, 
                presenterAction,
                typeof(InvalidRefactoringModelException));
            Assert.AreEqual(inputCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void Model_NoTargetAtSelection_TakesModule()
        {
            const string inputCode =
                @"
Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            var tdo = new RenameTestsDataObject("Class1", Selection.Home, newName: "RenamedClass1");
            var inputOutput = new RenameTestModuleDefinition("Class1")
            {
                Input = inputCode,
                Expected = inputCode,
                NewName = "RenamedClass1"
            };

            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameParameter_DoesNotAlterPrecompilerDirectives()
        {
            var tdo = new RenameTestsDataObject(selectedIdentifier: "arg1", newName: "arg2");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"#Const Bar = 42

#If False Then
Private Sub Goo(ByVal arg1 As String)
#ElseIf True Then
Private Sub Foo(ByVal ar|g1 As String)
#Else
Private Sub Foo(ByVal arg1 As String, arg2 As String)
#End If
End Sub",
                Expected =
                    @"#Const Bar = 42

#If False Then
Private Sub Goo(ByVal arg1 As String)
#ElseIf True Then
Private Sub Foo(ByVal arg2 As String)
#Else
Private Sub Foo(ByVal arg1 As String, arg2 As String)
#End If
End Sub"
            };
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameViewModel_IsValidName_ChangeCasingNotValid()
        {
            const string input =
                    @"Private Sub Foo()
    Dim val1 As Integer
End Sub";
            const string selected = "val1";
            const string newName = "Val1";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(input, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declaration = state.DeclarationFinder
                    .DeclarationsWithType(DeclarationType.Variable)
                    .First(d => d.IdentifierName.Equals(selected));
                var renameModel = new RenameModel(declaration);
                var messageBox = new Mock<IMessageBox>().Object;
                var renameViewModel = new RenameViewModel(state, renameModel, messageBox);
                renameViewModel.NewName = newName;
                Assert.IsFalse(renameViewModel.IsValidName); 
            }
        }


        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameClassModule_DoesNotChangeMeReferences()
        {
            const string newName = "RenamedClassModule";

            const string inputCode =
                @"Property Get Self() As IClassModule
    Set Self = Me
End Property";

            var tdo = new RenameTestsDataObject("Class1", DeclarationType.ClassModule, newName);
            var inputOutput = new RenameTestModuleDefinition("Class1")
            {
                Input = inputCode,
                Expected = inputCode,
                NewName =  newName
            };

            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }
        #endregion

        #region Test Setup

        private void PerformExpectedVersusActualRenameTests(RenameTestsDataObject tdo, params RenameTestModuleDefinition[] testModuleDefinitions)
        {
            InitializeTestDataObject(tdo, testModuleDefinitions);
            RunRenameRefactorScenario(tdo);
            CheckRenameRefactorTestResults(tdo);
        }

        private static void InitializeTestDataObject(RenameTestsDataObject tdo, params RenameTestModuleDefinition[] testModuleDefinitions)
        {
            if (tdo.RefactorParamType != RefactorParams.Declaration && tdo.RawSelection == null)
            {
                VerifyExactlyOneModuleHasASelection(testModuleDefinitions);
                DetermineSelectionFromFauxCursor(tdo, testModuleDefinitions);
            }

            tdo.ModuleTestSetupDefs.AddRange(testModuleDefinitions);

            if (tdo.NewName.Length == 0)
            {
                Assert.Inconclusive("NewName is blank");
            }
            if (tdo.RefactorParamType != RefactorParams.Declaration && !tdo.RawSelection.HasValue)
            {
                Assert.Inconclusive("A User 'Selection' has not been defined for the test");
            }

            tdo.PresenterAdjustmentAction = tdo.DoNotRename
                ? CaptureModel(tdo, model => model)
                : CaptureModel(tdo, AdjustName(tdo.NewName));

            tdo.VBE = tdo.VBE ?? BuildProject(tdo.ProjectName, tdo.ModuleTestSetupDefs, tdo.UseLibraries);
            tdo.AdditionalSetup?.Invoke(tdo);

            if (tdo.RefactorParamType != RefactorParams.Declaration)
            {
                CreateQualifiedSelectionForTestCase(tdo);
            }
        }

        private static Func<RenameModel, RenameModel> AdjustName(string newName)
        {
            return model =>
            {
                model.NewName = newName;
                return model;
            };
        }

        private static Func<RenameModel, RenameModel> CaptureModel(RenameTestsDataObject tdo, Func<RenameModel, RenameModel> presenterAction)
        {
            return model =>
            {
                tdo.Model = model;
                return presenterAction(model);
            };
        }

        private static void VerifyExactlyOneModuleHasASelection(IEnumerable<RenameTestModuleDefinition> testModuleDefinitions)
        {
            var cursorFound = false;
            foreach (var testModuleDefinition in testModuleDefinitions)
            {
                if (testModuleDefinition.InputWithFauxCursor.Equals(string.Empty))
                {
                    continue;
                }

                if (cursorFound)
                {
                    Assert.Inconclusive($"Found multiple selection cursors ('{FAUX_CURSOR}') in the test input");
                }

                cursorFound = true;
            }

            if (!cursorFound)
            {
                Assert.Inconclusive($"Unable to determine selected target using '{FAUX_CURSOR}' in test input");
            }
        }

        private static void DetermineSelectionFromFauxCursor(RenameTestsDataObject tdo, IEnumerable<RenameTestModuleDefinition> testModuleDefinitions)
        {
            var selectedTestModule = testModuleDefinitions.Single(testModuleDefinition =>
                testModuleDefinition.InputWithFauxCursor != string.Empty);

            tdo.SelectionModuleName = selectedTestModule.ModuleName;
            if (selectedTestModule.InputWithFauxCursor.Contains(RenameTests.FAUX_CURSOR))
            {
                var numCursors = selectedTestModule.InputWithFauxCursor.ToArray().Count(c => c.Equals(FAUX_CURSOR));
                if (numCursors != 1)
                {
                    Assert.Inconclusive($"{numCursors} found in FauxCursor input - only a single cursor is allowed.");
                }

                tdo.RawSelection = selectedTestModule.RenameSelection;

                if (!tdo.RawSelection.HasValue)
                {
                    Assert.Inconclusive($"Unable to set RawSelection field for test module {selectedTestModule.ModuleName}");
                }
            }
        }

        private void RunRenameRefactorScenario(RenameTestsDataObject tdo)
        {
            if (tdo.RefactorParamType == RefactorParams.Declaration)
            {
                tdo.ActualCode = RefactoredCode(tdo.VBE, tdo.TargetDeclarationName, tdo.TargetDeclarationType,
                    tdo.PresenterAdjustmentAction, tdo.ExpectedException);
            }
            else if (tdo.RefactorParamType == RefactorParams.QualifiedSelection)
            {
                tdo.ActualCode = RefactoredCode(tdo.VBE, tdo.SelectionModuleName, tdo.QualifiedSelection.Selection, tdo.PresenterAdjustmentAction, tdo.ExpectedException);
            }
            else
            {
                tdo.ActualCode = RefactoredCode(tdo.VBE, tdo.SelectionModuleName, tdo.QualifiedSelection.Selection, tdo.PresenterAdjustmentAction, tdo.ExpectedException, executeViaActiveSelection: true);
            }
        }

        private static void CheckRenameRefactorTestResults(RenameTestsDataObject tdo)
        {
            foreach (var inputOutput in tdo.ModuleTestSetupDefs)
            {
                if (inputOutput.CheckExpectedEqualsActual)
                {
                    var expected = inputOutput.Expected;
                    var actual = tdo.ActualCode[inputOutput.NewName];
                    Assert.AreEqual(expected, actual);
                }
            }
        }

        private static void CreateQualifiedSelectionForTestCase(RenameTestsDataObject tdo)
        {
            var component = RetrieveComponent(tdo, tdo.SelectionModuleName);
            if (tdo.RawSelection.HasValue)
            {
                tdo.QualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), tdo.RawSelection.Value);
                return;
            }
            Assert.Inconclusive($"Unable to find target '{FAUX_CURSOR}' in { tdo.SelectionModuleName} content.");
        }

        private static IVBE BuildProject(string projectName, IEnumerable<RenameTestModuleDefinition> testModuleDefinitions, bool useLibraries = false)
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(projectName, ProjectProtection.Unprotected);

            if (useLibraries)
            {
                enclosingProjectBuilder.AddReference(ReferenceLibrary.VBA);
                enclosingProjectBuilder.AddReference(ReferenceLibrary.Excel);
            }

            foreach (var testModuleDefinition in testModuleDefinitions)
            {
                if (testModuleDefinition.ModuleType == ComponentType.UserForm)
                {
                    var form = enclosingProjectBuilder.MockUserFormBuilder(testModuleDefinition.ModuleName, testModuleDefinition.Input);
                    if (!testModuleDefinition.ControlNames.Any())
                    {
                        Assert.Inconclusive("Test incorporates a UserForm but does not define any controls");
                    }
                    foreach (var control in testModuleDefinition.ControlNames)
                    {
                        form.AddControl(control);
                    }
                    (var component, var codeModule) = form.Build();
                    enclosingProjectBuilder.AddComponent(component, codeModule);
                }
                else
                {
                    var selection = testModuleDefinition.RenameSelection.HasValue 
                        ? testModuleDefinition.RenameSelection.Value 
                        : default;
                    enclosingProjectBuilder.AddComponent(testModuleDefinition.ModuleName, testModuleDefinition.ModuleType, testModuleDefinition.Input, selection);
                }
            }
            var project = enclosingProjectBuilder.Build();
            builder.AddProject(project);
            var vbe = builder.Build();
            return vbe.Object;
        }

        internal static IVBComponent RetrieveComponent(RenameTestsDataObject tdo, string componentName)
        {
            var vbProject = tdo.VBE.VBProjects.Single(item => item.Name == tdo.ProjectName);
            return vbProject.VBComponents.SingleOrDefault(item => item.Name == componentName);
        }

        internal enum RefactorParams
        {
            None,
            QualifiedSelection,
            Declaration
        }

        public void Assert_WarnsAboutRenameConflict_ConfirmationOutcomeHasExpectedResult(
            string inputCode,
            ComponentType componentType,
            string targetDeclarationName,
            DeclarationType targetDeclarationType,
            string newName,
            bool confirm,
            RefactoringDialogResult expectedDialogResult)
        {
            var inputOutput = new RenameTestModuleDefinition("Module1", componentType)
            {
                Input = inputCode
            };

            var vbe = BuildProject("TestProject", new[] {inputOutput});
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var target = state.DeclarationFinder
                    .DeclarationsWithType(targetDeclarationType)
                    .First(d => d.IdentifierName.Equals(targetDeclarationName));

                var renameModel = new RenameModel(target);
                var messageBoxMock = new Mock<IMessageBox>();
                messageBoxMock.Setup(m => m.ConfirmYesNo(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>()))
                    .Returns(() => confirm);

                var renameViewModel = new RenameViewModel(state, renameModel, messageBoxMock.Object)
                {
                    NewName = newName
                };

                if (expectedDialogResult == RefactoringDialogResult.Cancel)
                {
                    renameViewModel.OnWindowClosed += AssertDialogCancel;
                }
                else
                {
                    renameViewModel.OnWindowClosed += AssertDialogExecute;
                }

                renameViewModel.OkButtonCommand.Execute(null);

                messageBoxMock.Verify(m => m.ConfirmYesNo(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>()), Times.Once);
            }
        }

        private void AssertDialogExecute(object requestor, RefactoringDialogResult dialogResult)
        {
            Assert.AreEqual(RefactoringDialogResult.Execute, dialogResult);
        }

        private void AssertDialogCancel(object requestor, RefactoringDialogResult dialogResult)
        {
            Assert.AreEqual(RefactoringDialogResult.Cancel, dialogResult);
        }

        public void Assert_WarnsAboutTargetMove_DenialOfConfirmationLeadsToAbort(RenameModel model, bool confirm)
        {
            if (model.Target.Equals(model.InitialTarget))
            {
                Assert.Inconclusive("The actual target and the initial one agree.");
            }

            var dialogMock = new Mock<IRefactoringDialog<RenameModel, IRefactoringView<RenameModel>, IRefactoringViewModel<RenameModel>>>();
            dialogMock.Setup(m => m.ShowDialog())
                .Callback(() => throw new DialogSuccessfullyEnteredTestException());
            var dialogFactoryMock = new Mock<IRefactoringDialogFactory>();
            dialogFactoryMock.Setup(m => m.CreateDialog
                <
                    RenameModel,
                    IRefactoringView<RenameModel>,
                    IRefactoringViewModel<RenameModel>,
                    IRefactoringDialog<RenameModel, IRefactoringView<RenameModel>, IRefactoringViewModel<RenameModel>>
                >(
                    It.IsAny<DialogData>(),
                    It.IsAny<RenameModel>(),
                    It.IsAny<IRefactoringView<RenameModel>>(),
                    It.IsAny<IRefactoringViewModel<RenameModel>>()
                ))
                .Returns(
                    (
                        DialogData dialogData, 
                        RenameModel renameModel, 
                        IRefactoringView<RenameModel> view, 
                        IRefactoringViewModel<RenameModel>  viewModel
                    ) =>
                    {
                        dialogMock.SetupGet(m => m.Model).Returns(() => renameModel);
                        return dialogMock.Object;
                    }
                );

            var messageBoxMock = new Mock<IMessageBox>();
            messageBoxMock.Setup(m => m.ConfirmYesNo(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>()))
                .Returns(confirm);

            var presenter = new RenamePresenter(model, dialogFactoryMock.Object, messageBoxMock.Object);

            if (confirm)
            {
                Assert.Throws<DialogSuccessfullyEnteredTestException>(() => presenter.Show());
            }
            else
            {
                Assert.Throws<RefactoringAbortedException>(() => presenter.Show());
            }
            
            messageBoxMock.Verify(m => m.ConfirmYesNo(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>()), Times.Once);
        }

        protected override IRefactoring TestRefactoring(
            IRewritingManager rewritingManager, 
            RubberduckParserState state,
            RefactoringUserInteraction<IRenamePresenter, RenameModel> userInteraction, 
            ISelectionService selectionService)
        {
            var selectedDeclarationService = new SelectedDeclarationProvider(selectionService, state);
            var componentRename = new RenameComponentOrProjectRefactoringAction(state, state?.ProjectsProvider, state, rewritingManager);
            var otherRename = new RenameCodeDefinedIdentifierRefactoringAction(state, state?.ProjectsProvider, rewritingManager);
            var baseRefactoring = new RenameRefactoringAction(componentRename, otherRename);
            return new RenameRefactoring(baseRefactoring, userInteraction, state, state?.ProjectsProvider, selectionService, selectedDeclarationService);
        }

        #endregion

        private class DialogSuccessfullyEnteredTestException : Exception
        {}
    }
}
