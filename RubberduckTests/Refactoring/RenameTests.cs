using System;
using System.Linq;
using System.Windows.Forms;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using System.Collections.Generic;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Refactorings.Rename;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class RenameTests
    {
        private const string FAUX_CURSOR = "|";

        #region Rename Variable Tests

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameVariable()
        {
            var tdo = new RenameTestsDataObject(selection: "val1", newName: "val2");
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
            var tdo = new RenameTestsDataObject(selection: "val1", newName: "val2");
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

        #endregion
        #region Rename Parameter Tests

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameParameter()
        {
            var tdo = new RenameTestsDataObject(selection: "arg1", newName: "arg2");
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
            var tdo = new RenameTestsDataObject(selection: "arg3", newName: "arg2");
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
            var tdo = new RenameTestsDataObject(selection: "arg1", newName: "arg2");
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
            var tdo = new RenameTestsDataObject(selection: "index", newName: "renamed");
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
        public void RenameRefactoring_RenameLastPropertyParameter_UpdatesAllRelatedParameters()
        {
            var tdo = new RenameTestsDataObject(selection: "value", newName: "renamed");
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
        public void RenameRefactoring_RenameLastPropertyParameter_UpdatesRelatedParametersWithSameName()
        {
            var tdo = new RenameTestsDataObject(selection: "value", newName: "renamed");
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
        public void RenameRefactoring_RenameSub_ConflictingNames_Reject()
        {
            var tdo = new RenameTestsDataObject(selection: "Foo", newName: "Goo");
            var inputOutput = new RenameTestModuleDefinition("Module1", ComponentType.StandardModule)
            {
                Input =
                    @"Private Sub Fo|o()
    Dim Goo As Integer
End Sub",
                Expected =
                    @"Private Sub Foo()
    Dim Goo As Integer
End Sub"
            };
            tdo.MsgBoxReturn = DialogResult.No;
            PerformExpectedVersusActualRenameTests(tdo, inputOutput);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameSub_ConflictingNames_Accept()
        {
            var tdo = new RenameTestsDataObject(selection: "Foo", newName: "Goo");
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
            var tdo = new RenameTestsDataObject(selection: "Foo", newName: "Hoo");
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
            var tdo = new RenameTestsDataObject(selection: "Foo", newName: "Goo");
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
            var tdo = new RenameTestsDataObject(selection: "Foo", newName: "Goo");
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
            var tdo = new RenameTestsDataObject(selection: "Foo", newName: "Hoo");
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
            var tdo = new RenameTestsDataObject(selection: "Foo", newName: "Hoo");
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

        #endregion
        #region Rename Control Tests
        //All RenameControl tests are ignored because control renames depend on access to
        //Non-UserDefined declarations in the DeclarationFinder.  So, the control rename scenarios
        //below can only be tested if implemented (and tested) within Excel.  

        [Test, Ignore("")]
        public void RenameRefactoring_RenameControlFromEventHandler()
        {
            var tdo = new RenameTestsDataObject(selection: "cmdBtn1", newName: "cmdBigButton");
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
        public void RenameRefactoring_RenameControlFromEventHandlerNameCollision()
        {
            var tdo = new RenameTestsDataObject(selection: "cmdBtn1", newName: "cmdBigButton");
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
        public void RenameRefactoring_RenameControlRenameInReference()
        {
            var tdo = new RenameTestsDataObject(selection: "cmdBtn1", newName: "cmdBigButton");
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

            tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()), Times.Never);
        }

        [Test, Ignore("")]
        public void RenameRefactoring_RenameControlFromEventHandlerReference()
        {
            var tdo = new RenameTestsDataObject(selection: "cmdBtn1", newName: "cmdBigButton");
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

            tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.OK, It.IsAny<MessageBoxIcon>()), Times.Never);
        }

        [Test, Ignore("")]
        public void RenameRefactoring_RenameControlHandlesUnderscoresInNewName()
        {
            var tdo = new RenameTestsDataObject(selection: "bigButton_ClickAgain", newName: "bigButton_ClickAgain_AndAgain");
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
        public void RenameRefactoring_RenameControlSimilarNames()
        {
            var tdo = new RenameTestsDataObject(selection: "bigButton", newName: "smallButton");
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

        #endregion
        #region Rename Event Tests

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameEvent()
        {
            var tdo = new RenameTestsDataObject(selection: "Foo", newName: "Goo");
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
            var tdo = new RenameTestsDataObject(selection: "Foo", newName: "Goo");
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
            var tdo = new RenameTestsDataObject(selection: "abc_Foo", newName: "abc_Goo");
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
            var tdo = new RenameTestsDataObject(selection: "def_Foo", newName: "def_Goo");
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
            var tdo = new RenameTestsDataObject(selection: "Foo", newName: "Goo");
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
            var tdo = new RenameTestsDataObject(selection: "MyEvent", newName: "YourEvent");
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
            var tdo = new RenameTestsDataObject(selection: "XLEvents", newName: "NewEventImpl");
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
            var tdo = new RenameTestsDataObject(selection: "abc", newName: "def");
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
            var tdo = new RenameTestsDataObject(selection: "MyEvent", newName: "YourEvent_withUnderscore");
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
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameEventFromUsage()
        {
            var tdo = new RenameTestsDataObject(selection: "MyEvent", newName: "YourEvent");
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
            var tdo = new RenameTestsDataObject(selection: "DoSomething", newName: "DoNothing");
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

            tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()), Times.Never);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameInterfaceMemberDuplicateMemberInOtherInterface()
        {
            var tdo = new RenameTestsDataObject(selection: "DoSomething", newName: "DoNothing");
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
            var tdo = new RenameTestsDataObject(selection: "DoSomething", newName: "DoNothing");
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
            var tdo = new RenameTestsDataObject(selection: "DoSomething", newName: "DoNothing");
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

            tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()), Times.Once);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameInterfaceFromMemberProperty()
        {
            var tdo = new RenameTestsDataObject(selection: "Something", newName: "Nothing");
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

            tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()), Times.Once);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameInterfaceNoImplementers()
        {
            var tdo = new RenameTestsDataObject(selection: "DoSomething", newName: "DoNothing");
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

            tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()), Times.Never);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameInterfaceFromReference()
        {
            var tdo = new RenameTestsDataObject(selection: "DoSomething", newName: "DoNothing");
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
            var tdo = new RenameTestsDataObject(selection: "DoSomething", newName: "DoNothing");
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
            var tdo = new RenameTestsDataObject(selection: "DoSomething", newName: "DoNothing");
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

            tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()), Times.Once);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameInterface_RejectPrompt()
        {
            var tdo = new RenameTestsDataObject(selection: "DoSomething", newName: "DoNothing");
            var inputOutput1 = new RenameTestModuleDefinition("Class1")
            {
                Input =
                    @"Implements IClass1

Private Sub ICla|ss1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub"
            };
            inputOutput1.Expected = inputOutput1.Input.Replace(FAUX_CURSOR, "");

            var inputOutput2 = new RenameTestModuleDefinition("IClass1")
            {
                Input =
                    @"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub"
            };
            inputOutput2.Expected = inputOutput2.Input;

            tdo.MsgBoxReturn = DialogResult.No;
            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutput2);

            tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()), Times.Once);
        }

        #endregion
        #region Rename CodeModule Tests

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameModuleFromImplementsStmt()
        {
            var tdo = new RenameTestsDataObject(selection: "IClass1", newName: "INewClass");
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
            var tdo = new RenameTestsDataObject(selection: "CTestClass", newName: "CMyTestClass");
            var inputOutput1 = new RenameTestModuleDefinition("CTestClass")
            {
                Input =
                    @"
Sub Foo()
End Sub"
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

            PerformExpectedVersusActualRenameTests(tdo, inputOutput1, inputOutput2);

            var component = RetrieveComponent(tdo, inputOutput1.ModuleName);
            Assert.AreSame(tdo.NewName, component.CodeModule.Name);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameCodeModule()
        {
            const string newName = "RenameModule";

            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal a As Integer, ByVal b As String)
End Sub";

            var selection = new Selection(3, 27, 3, 27);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, "Class1", ComponentType.ClassModule, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var msgbox = new Mock<IMessageBox>();
                msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                    .Returns(DialogResult.Yes);

                var vbeWrapper = vbe.Object;
                var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = newName };
                model.Target = model.Declarations.FirstOrDefault(i => i.DeclarationType == DeclarationType.ClassModule && i.IdentifierName == "Class1");

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
                refactoring.Refactor(model.Target);

                Assert.AreSame(newName, component.CodeModule.Name);
            }
        }

        #endregion
        #region Rename Project Tests

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameProject()
        {
            const string oldName = "TestProject1";
            const string newName = "Renamed";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder(oldName, ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, string.Empty)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var msgbox = new Mock<IMessageBox>();
                msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                    .Returns(DialogResult.Yes);

                var vbeWrapper = vbe.Object;
                var model = new RenameModel(vbeWrapper, state, default(QualifiedSelection)) { NewName = newName };
                model.Target = model.Declarations.First(i => i.DeclarationType == DeclarationType.Project && i.IsUserDefined);

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
                refactoring.Refactor(model.Target);

                Assert.AreEqual(newName, vbe.Object.VBProjects[0].Name);
            }
        }

        #endregion
        #region Rename Enumeration Tests

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameEnumeration()
        {
            var tdo = new RenameTestsDataObject(selection: "FruitType", newName: "Fruits");
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

            tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()), Times.Never);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameEnumerationMember()
        {
            var tdo = new RenameTestsDataObject(selection: "Apple", newName: "CranApple");
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

            tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()), Times.Never);
        }

        #endregion
        #region Rename Label Tests
        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameLabel()
        {
            var tdo = new RenameTestsDataObject(selection: "EH", newName: "ErrorHandler");
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

            tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()), Times.Never);
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
                var tdo = new RenameTestsDataObject(selection: "Foo", newName: "Goo");
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

                tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()), Times.Never);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var codePaneMock = new Mock<ICodePane>();
                codePaneMock.Setup(c => c.CodeModule).Returns(component.CodeModule);
                codePaneMock.Setup(c => c.Selection);
                vbe.Setup(v => v.ActiveCodePane).Returns(codePaneMock.Object);

                var vbeWrapper = vbe.Object;
                var factory = new RenamePresenterFactory(vbeWrapper, null, state);

                var refactoring = new RenameRefactoring(vbeWrapper, factory, null, state);
                refactoring.Refactor();

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(inputCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void Presenter_TargetIsNull()
        {
            const string inputCode =
                @"
Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var codePaneMock = new Mock<ICodePane>();
                codePaneMock.Setup(c => c.CodeModule).Returns(component.CodeModule);
                codePaneMock.Setup(c => c.Selection);
                vbe.Setup(v => v.ActiveCodePane).Returns(codePaneMock.Object);

                var vbeWrapper = vbe.Object;
                var factory = new RenamePresenterFactory(vbeWrapper, null, state);

                var presenter = factory.Create();

                Assert.AreEqual(null, presenter.Show());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void Factory_SelectionIsNull()
        {
            const string inputCode =
                @"Private Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var codePaneMock = new Mock<ICodePane>();
                codePaneMock.Setup(c => c.CodeModule).Returns(component.CodeModule);
                codePaneMock.Setup(c => c.Selection);
                vbe.Setup(v => v.ActiveCodePane).Returns(codePaneMock.Object);

                var vbeWrapper = vbe.Object;
                var factory = new RenamePresenterFactory(vbeWrapper, null, state);

                var presenter = factory.Create();
                Assert.AreEqual(null, presenter.Show());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameParameter_DoesNotAlterPrecompilerDirectives()
        {
            var tdo = new RenameTestsDataObject(selection: "arg1", newName: "arg2");
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
            var tdo = new RenameTestsDataObject(selection: "val1", newName: "Val1");
            var inputOutput = new RenameTestModuleDefinition("TestClass")
            {
                Input =
                    @"Private Sub Foo()
    Dim va|l1 As Integer
End Sub",
                CheckExpectedEqualsActual = false
            };
            InitializeTestDataObject(tdo, inputOutput);
            var renameViewModel = new RenameViewModel(tdo.RenameModel.State);
            renameViewModel.Target = tdo.RenameModel.Target;
            renameViewModel.NewName = tdo.NewName;
            Assert.IsFalse(renameViewModel.IsValidName);
        }


        [Test]
        [Category("Refactorings")]
        [Category("Rename")]
        public void RenameRefactoring_RenameClassModule_DoesNotChangeMeReferences()
        {
            const string newName = "RenamedClassModule";

            //Input
            const string inputCode =
                @"Property Get Self() As IClassModule
    Set Self = Me
End Property";

            var selection = new Selection(3, 27, 3, 27);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, "ClassModule1", ComponentType.ClassModule, out component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var msgbox = new Mock<IMessageBox>();
                msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                    .Returns(DialogResult.Yes);

                var vbeWrapper = vbe.Object;
                var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = newName };
                model.Target = model.Declarations.FirstOrDefault(i => i.DeclarationType == DeclarationType.ClassModule && i.IdentifierName == "ClassModule1");

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
                refactoring.Refactor(model.Target);

                Assert.AreSame(newName, component.CodeModule.Name);
                Assert.AreEqual(inputCode, component.CodeModule.GetLines(0, component.CodeModule.CountOfLines));
            }

        }
        #endregion

        #region Test Execution

        private static void PerformExpectedVersusActualRenameTests(RenameTestsDataObject tdo
            , RenameTestModuleDefinition? inputOutput1
            , RenameTestModuleDefinition? inputOutput2 = null
            , RenameTestModuleDefinition? inputOutput3 = null
            , RenameTestModuleDefinition? inputOutput4 = null)
        {
            try
            {
                InitializeTestDataObject(tdo, inputOutput1, inputOutput2, inputOutput3, inputOutput4);
                RunRenameRefactorScenario(tdo);
                CheckRenameRefactorTestResults(tdo);
            }
            finally
            {
                tdo.ParserState?.Dispose();
            }
        }

        private static void InitializeTestDataObject(RenameTestsDataObject tdo
            , RenameTestModuleDefinition? inputOutput1
            , RenameTestModuleDefinition? inputOutput2 = null
            , RenameTestModuleDefinition? inputOutput3 = null
            , RenameTestModuleDefinition? inputOutput4 = null)
        {
            var renameTMDs = new List<RenameTestModuleDefinition>();
            bool cursorFound = false;
            foreach (var io in new[] { inputOutput1, inputOutput2, inputOutput3, inputOutput4 })
            {
                if (io.HasValue)
                {
                    var renameTMD = io.Value;
                    if (renameTMD.Input.Contains(FAUX_CURSOR))
                    {
                        if (cursorFound) { Assert.Inconclusive($"Found multiple selection cursors ('{FAUX_CURSOR}') in the test input"); }
                        cursorFound = true;
                        renameTMD.Input_WithFauxCursor = renameTMD.Input;
                        renameTMD.Input = renameTMD.Input.Replace(FAUX_CURSOR, "");
                    }
                    renameTMDs.Add(renameTMD);
                }
            }

            if (!cursorFound)
            {
                Assert.Inconclusive($"Unable to determine selected target using '{FAUX_CURSOR}' in test input");
            }

            renameTMDs.ForEach(rtmd => AddTestModuleDefinition(tdo, rtmd));

            if (tdo.NewName.Length == 0)
            {
                Assert.Inconclusive("NewName is blank");
            }
            if (!tdo.RawSelection.HasValue)
            {
                Assert.Inconclusive("A User 'Selection' has not been defined for the test");
            }

            tdo.MsgBox = new Mock<IMessageBox>();
            tdo.MsgBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                .Returns(tdo.MsgBoxReturn);

            tdo.VBE = tdo.VBE ?? BuildProject(tdo.ProjectName, tdo.ModuleTestSetupDefs);
            tdo.ParserState = MockParser.CreateAndParse(tdo.VBE);

            CreateQualifiedSelectionForTestCase(tdo);
            tdo.RenameModel = new RenameModel(tdo.VBE, tdo.ParserState, tdo.QualifiedSelection) { NewName = tdo.NewName };
            Assert.IsTrue(tdo.RenameModel.Target.IdentifierName.Contains(tdo.OriginalName)
                , $"Target aquired ({tdo.RenameModel.Target.IdentifierName} does not equal name specified ({tdo.OriginalName}) in the test");

            var factory = SetupFactory(tdo.RenameModel);

            tdo.RenameRefactoringUnderTest = new RenameRefactoring(tdo.VBE, factory.Object, tdo.MsgBox.Object, tdo.ParserState);
        }

        private static void AddTestModuleDefinition(RenameTestsDataObject tdo, RenameTestModuleDefinition inputOutput)
        {
            if (inputOutput.RenameSelection.HasValue)
            {
                tdo.SelectionModuleName = inputOutput.ModuleName;
                tdo.RawSelection = inputOutput.RenameSelection;
            }
            else if (inputOutput.Input_WithFauxCursor.Length > 0)
            {
                tdo.SelectionModuleName = inputOutput.ModuleName;
                if (inputOutput.Input_WithFauxCursor.Contains(FAUX_CURSOR))
                {
                    var numCursors = inputOutput.Input_WithFauxCursor.ToArray().Where(c => c.Equals('|')).Count();
                    if (numCursors != 1)
                    {
                        Assert.Inconclusive($"{numCursors} found in FauxCursor input - only a single cursor is allowed.");
                    }
                    tdo.RawSelection = FindSelection(inputOutput.Input_WithFauxCursor, tdo.OriginalName);
                    if (!tdo.RawSelection.HasValue)
                    {
                        Assert.Inconclusive($"Unable to set RawSelection field for test module {inputOutput.ModuleName}");
                    }
                }
            }
            tdo.ModuleTestSetupDefs.Add(inputOutput);
        }

        private static Selection? FindSelection(string input, string originalName)
        {
            var lines = input.Split(new[] { "\r\n" }, StringSplitOptions.None);
            for (var idx = 0; idx < lines.Count(); idx++)
            {
                var testLine = lines[idx];
                if (testLine.Contains(FAUX_CURSOR))
                {
                    var fauxCursorLine = idx + 1;
                    var fauxCursorColumn = testLine.IndexOf(FAUX_CURSOR);
                    if (!testLine.Replace(FAUX_CURSOR, "").Contains(originalName))
                    {
                        Assert.Inconclusive($"Module line with faux cursor does not contain target '{originalName}'");
                    }

                    int fauxCursorColumnOffset = 0;
                    if (testLine.StartsWith($"{FAUX_CURSOR}") || testLine.Contains($" {FAUX_CURSOR}"))
                    {
                        fauxCursorColumnOffset = 1;
                    }
                    else if (testLine.EndsWith($"{FAUX_CURSOR}") || testLine.Contains($"{FAUX_CURSOR} "))
                    {
                        fauxCursorColumnOffset = -1;
                    }

                    return new Selection(fauxCursorLine, fauxCursorColumn + fauxCursorColumnOffset);
                }
            }
            return null;
        }

        private static void RunRenameRefactorScenario(RenameTestsDataObject tdo)
        {
            if (tdo.RefactorParamType == RefactorParams.Declaration)
            {
                tdo.RenameRefactoringUnderTest.Refactor(tdo.RenameModel.Target);
            }
            else if (tdo.RefactorParamType == RefactorParams.QualifiedSelection)
            {
                tdo.RenameRefactoringUnderTest.Refactor(tdo.QualifiedSelection);
            }
            else
            {
                tdo.RenameRefactoringUnderTest.Refactor();
            }
        }

        private static void CheckRenameRefactorTestResults(RenameTestsDataObject tdo)
        {
            foreach (var inputOutput in tdo.ModuleTestSetupDefs)
            {
                if (inputOutput.CheckExpectedEqualsActual)
                {
                    var rewriter = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, inputOutput.ModuleName).CodeModule.Parent);
                    Assert.AreEqual(inputOutput.Expected, rewriter.GetText());
                }
            }
        }

        private static Mock<IRefactoringPresenterFactory<IRenamePresenter>> SetupFactory(RenameModel model)
        {
            var presenter = new Mock<IRenamePresenter>();
            presenter.Setup(p => p.Model).Returns(model);
            presenter.Setup(p => p.Show()).Returns(model);
            presenter.Setup(p => p.Show(It.IsAny<Declaration>())).Returns(model);

            var factory = new Mock<IRefactoringPresenterFactory<IRenamePresenter>>();
            factory.Setup(f => f.Create()).Returns(presenter.Object);
            return factory;
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

        private static IVBE BuildProject(string projectName, List<RenameTestModuleDefinition> testComponents)
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(projectName, ProjectProtection.Unprotected);

            foreach (var comp in testComponents)
            {
                if (comp.ModuleType == ComponentType.UserForm)
                {
                    var form = enclosingProjectBuilder.MockUserFormBuilder(comp.ModuleName, comp.Input);
                    if (!comp.ControlNames.Any())
                    {
                        Assert.Inconclusive("Test incorporates a UserForm but does not define any controls");
                    }
                    foreach (var control in comp.ControlNames)
                    {
                        form.AddControl(control);
                    }
                    enclosingProjectBuilder.AddComponent(form.Build());
                }
                else
                {
                    enclosingProjectBuilder.AddComponent(comp.ModuleName, comp.ModuleType, comp.Input);
                }
            }

            builder.AddProject(enclosingProjectBuilder.Build());
            return builder.Build().Object;
        }

        private static IVBComponent RetrieveComponent(RenameTestsDataObject tdo, string componentName)
        {
            var vbProject = tdo.VBE.VBProjects.Single(item => item.Name == tdo.ProjectName);
            return vbProject.VBComponents.SingleOrDefault(item => item.Name == componentName);
        }

        internal struct RenameTestModuleDefinition
        {
            public string Input_WithFauxCursor;
            public string Input;
            public string Expected;
            public string ModuleName;
            public ComponentType ModuleType;
            public bool CheckExpectedEqualsActual;
            public List<string> ControlNames;
            public string NewName;
            public Selection? RenameSelection;

            public RenameTestModuleDefinition(string moduleName, ComponentType moduleType = ComponentType.ClassModule)
            {
                Input_WithFauxCursor = "";
                Expected = "";
                Input = "";
                ModuleName = moduleName;
                ModuleType = moduleType;
                CheckExpectedEqualsActual = true;
                ControlNames = new List<String>();
                NewName = "";
                RenameSelection = null;
            }
        }

        internal enum RefactorParams
        {
            None,
            QualifiedSelection,
            Declaration
        };

        internal class RenameTestsDataObject
        {
            public RenameTestsDataObject(string selection, string newName)
            {
                ProjectName = "TestProject";
                MsgBoxReturn = DialogResult.Yes;
                RefactorParamType = RefactorParams.QualifiedSelection;
                RawSelection = null;
                NewName = newName;
                OriginalName = selection;
                ModuleTestSetupDefs = new List<RenameTestModuleDefinition>();
                RenameRefactoringUnderTest = null;
            }

            public IVBE VBE { get; set; }
            public RubberduckParserState ParserState { get; set; }
            public string ProjectName { get; set; }
            public string NewName { get; set; }
            public string SelectionModuleName { get; set; }
            public QualifiedSelection QualifiedSelection { get; set; }
            public RenameModel RenameModel { get; set; }
            public Mock<IMessageBox> MsgBox { get; set; }
            public DialogResult MsgBoxReturn { get; set; }
            public RefactorParams RefactorParamType { get; set; }
            public Selection? RawSelection { get; set; }
            public List<RenameTestModuleDefinition> ModuleTestSetupDefs { get; set; }
            public string OriginalName { get; set; }
            public RenameRefactoring RenameRefactoringUnderTest { get; set; }
        }
        #endregion
    }
}
