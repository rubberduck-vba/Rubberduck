using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;

namespace RubberduckTests.Symbols
{
    [TestFixture]
    public class SelectedDeclarationProviderTests
    {

        [Test]
        [Category("Resolver")]
        public void FiendishlyAmbiguousNameSelectsSmallestScopedDeclaration()
        {
            var code =
            @"
Option Explicit

Public Sub foo()
    Dim foo As Long
    foo = 42
    Debug.Print foo
End Sub
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("foo", ProjectProtection.Unprotected)
                .AddComponent("foo", ComponentType.StandardModule, code, new Selection(6, 6))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Variable, "foo");

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void AmbiguousNameSelectsSmallestScopedDeclaration()
        {
            var code =
@"
Option Explicit

Public Sub foo()
    Dim foo As Long
    foo = 42
    Debug.Print foo
End Sub
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(6, 6))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Variable, "foo");

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void AmbiguousNameSelectsParameterOverProperty()
        {
            var code =
                @"
Option Explicit

Public Property Get Item()
    Item = 12
End Property

Public Property Let Item(ByVal Item As Variant)
    DoSomething Item
End Property

Private Sub DoSomething(ByVal value As Variant)
    Debug.Print value
End Sub
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(9, 18))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Parameter, "Item");

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void AmbiguousNameSelectsParameterOverSub()
        {
            var code =
                @"
Option Explicit

Public Sub foo(ByVal foo As Bookmarks)
    Dim bar As Bookmark
    For Each bar In foo
        Debug.Print bar.Name
    Next
End Sub
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(6, 22))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Parameter, "foo");

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void SameNameForProjectAndClass_ScopedDeclaration_ProjectSelection()
        {
            var refEditClass = @"
Option Explicit

Private ValueField As Variant

Public Property Get Value()
  Value = ValueField
End Property

Public Property Let Value(Value As Variant)
  ValueField = Value
End Property";

            var code =
                @"
Option Explicit

Public Sub foo()
    Dim myEdit As RefEdit.RefEdit
    Set myEdit = New RefEdit.RefEdit

    myEdit.Value = ""abc""
End Sub
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("RefEdit", ProjectProtection.Unprotected)
                .AddComponent("RefEdit", ComponentType.ClassModule, refEditClass)
                .AddComponent("Test", ComponentType.StandardModule, code, new Selection(6, 23))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Project, "RefEdit");

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void SameNameForProjectAndClass_ScopedDeclaration_ClassSelection()
        {
            var refEditClass = @"
Option Explicit

Private ValueField As Variant

Public Property Get Value()
  Value = ValueField
End Property

Public Property Let Value(Value As Variant)
  ValueField = Value
End Property";

            var code =
                @"
Option Explicit

Public Sub foo()
    Dim myEdit As RefEdit.RefEdit
    Set myEdit = New RefEdit.RefEdit

    myEdit.Value = ""abc""
End Sub
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("RefEdit", ProjectProtection.Unprotected)
                .AddComponent("RefEdit", ComponentType.ClassModule, refEditClass)
                .AddComponent("Test", ComponentType.StandardModule, code, new Selection(6, 31))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.ClassModule, "RefEdit");

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void Resolve_RecursiveCall_AssignmentLHS()
        {
            var sillyClass = @"
Option Explicit

Public Function Class1() As Class1
    Set Class1 = Me
End Function";

            var code =
                @"
Option Explicit

Public Function Class1(this As Class1) As Class1
    Set this = New Class1
    
    Set Class1 = Class1(this)
End Function";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, sillyClass)
                .AddComponent("Test", ComponentType.StandardModule, code, new Selection(7, 10))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Function, "Class1", "TestProject.Test");

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void Resolve_RecursiveCall_AssignmentRHS()
        {
            var sillyClass = @"
Option Explicit

Public Function Class1() As Class1
    Set Class1 = Me
End Function";

            var code =
                @"
Option Explicit

Public Function Class1(this As Class1) As Class1
    Set this = New Class1
    
    Set Class1 = Class1(this)
End Function";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, sillyClass)
                .AddComponent("Test", ComponentType.StandardModule, code, new Selection(7, 19))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Function, "Class1", "TestProject.Test");

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void Resolve_RecursiveCall_PropertyAssignment_PropertyLetAccessor()
        {
            var sillyClass = @"
Option Explicit

Public Property Get Class1() As Class1
    Set Class1 = Me
End Property

Public Property Let Class1(Class1 As Class1)
    Set Class1 = Class1
End Property";

            var code =
                @"
Option Explicit

Public Function Class1(this As Class1) As Class1
    Set this = New Class1
    
    Set Class1 = this.Class1(this)
End Function";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Test", ComponentType.StandardModule, code)
                .AddComponent("Class1", ComponentType.ClassModule, sillyClass, new Selection(8, 22))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.PropertyLet, "Class1");

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void Resolve_RecursiveCall_PropertyAssignment_Parameter()
        {
            var sillyClass = @"
Option Explicit

Public Property Get Class1() As Class1
    Set Class1 = Me
End Property

Public Property Let Class1(Class1 As Class1)
    Set Class1 = Class1
End Property";

            var code =
                @"
Option Explicit

Public Function Class1(this As Class1) As Class1
    Set this = New Class1
    
    Set Class1 = this.Class1(this)
End Function";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Test", ComponentType.StandardModule, code)
                .AddComponent("Class1", ComponentType.ClassModule, sillyClass, new Selection(8, 29))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Parameter, "Class1");

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void Resolve_RecursiveCall_PropertyAssignment_AsType()
        {
            var sillyClass = @"
Option Explicit

Public Property Get Class1() As Class1
    Set Class1 = Me
End Property

Public Property Let Class1(Class1 As Class1)
    Set Class1 = Class1
End Property";

            var code =
                @"
Option Explicit

Public Function Class1(this As Class1) As Class1
    Set this = New Class1
    
    Set Class1 = this.Class1(this)
End Function";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Test", ComponentType.StandardModule, code)
                .AddComponent("Class1", ComponentType.ClassModule, sillyClass, new Selection(8, 39))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.ClassModule, "Class1");

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void Resolve_RecursiveCall_PropertyAssignment_ParameterInBody()
        {
            var sillyClass = @"
Option Explicit

Public Property Get Class1() As Class1
    Set Class1 = Me
End Property

Public Property Let Class1(Class1 As Class1)
    Set Class1 = Class1
End Property";

            var code =
                @"
Option Explicit

Public Function Class1(this As Class1) As Class1
    Set this = New Class1
    
    Set Class1 = this.Class1(this)
End Function";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Test", ComponentType.StandardModule, code)
                .AddComponent("Class1", ComponentType.ClassModule, sillyClass, new Selection(9, 19))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Parameter, "Class1");

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void Resolve_RecursiveCall_PropertyAssignment_ParameterInAssignment()
        {
            // The assignment of the Property Let actually modifies the parameter, not the property getter
            // Therefore, we expect a parameter as the target of assignment. It is not actually recursive
            // though it may look like one.
            var sillyClass = @"
Option Explicit

Public Property Get Class1() As Class1
    Set Class1 = Me
End Property

Public Property Let Class1(Class1 As Class1)
    Set Class1 = Class1
End Property";

            var code =
                @"
Option Explicit

Public Function Class1(this As Class1) As Class1
    Set this = New Class1
    
    Set Class1 = this.Class1(this)
End Function";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Test", ComponentType.StandardModule, code)
                .AddComponent("Class1", ComponentType.ClassModule, sillyClass, new Selection(9, 10))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Parameter, "Class1");

            Assert.AreEqual(expected, actual);
        }

        [Category("Resolver")]
        [Test]
        public void Identify_NamedParameter_Parameter()
        {
            const string code = @"
Public Function Foo(Item As String) As Boolean
    MsgBox (Item)
End Function

Public Sub DoIt()
    Dim Result As Boolean
    Dim Item As String
    
    Item = ""abc""
    Result = Foo(Item:=Item)
End Sub";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(11, 19))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Parameter, "Item");

            Assert.AreEqual(expected, actual);
        }

        [Category("Resolver")]
        [Test]
        public void Identify_NamedParameter_LocalVariable()
        {
            const string code = @"
Public Function Foo(Item As String) As Boolean
    MsgBox (Item)
End Function

Public Sub DoIt()
    Dim Result As Boolean
    Dim Item As String
    
    Item = ""abc""
    Result = Foo(Item:=Item)
End Sub";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(11, 25))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Variable, "Item");

            Assert.AreEqual(expected, actual);
        }

        [Category("Resolver")]
        [Test]
        public void Identify_NamedParameter_Parameter_FromExcel()
        {
            const string code = @"
Public Sub DoIt()
    Dim sht As WorkSheet

    sht.Paste Link:=True
End Sub";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(5, 16))
                .AddReference(ReferenceLibrary.Excel)
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Parameter, "Link", "EXCEL.EXE;Excel._Worksheet.Paste");

            Assert.AreEqual(expected, actual);
        }

        [Category("Resolver")]
        [Test]
        public void SelectionInsideConstantDeclaration_ConstantSelected()
        {
            const string code = @"
Private Const myConst As Long = 42

Private myModuleVariable As Long


Public Sub DoIt()
    Dim myLocalVariable As Long

    myModuleVariable = myLocalVariable
End Sub";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(2, 2))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Constant, "myConst");

            Assert.AreEqual(expected, actual);
        }

        [Category("Resolver")]
        [Test]
        [TestCase(4, 2, "myModuleVariable")]
        [TestCase(8, 6, "myLocalVariable")]
        public void SelectionInsideVariableDeclaration_VariableSelected(int selectedLine, int selectedColumn, string expectedVariableName)
        {
            const string code = @"
Private Const myConst As Long = 42

Private myModuleVariable As Long


Public Sub DoIt()
    Dim myLocalVariable As Long

    myModuleVariable = myLocalVariable
End Sub";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(selectedLine, selectedColumn))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Variable, expectedVariableName);

            Assert.AreEqual(expected, actual);
        }

        [Category("Resolver")]
        [Test]
        public void SelectionInsideVariableDeclarationAtStartOfModule_VariableSelected()
        {
            const string code = @"Private myModuleVariable As Long


Public Sub DoIt()
    Dim myLocalVariable As Long

    myModuleVariable = myLocalVariable
End Sub";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Variable, "myModuleVariable");

            Assert.AreEqual(expected, actual);
        }

        [Category("Resolver")]
        [Test]
        public void SelectionInsideModuleBodyElementAndOnNothingElse_ModuleBodyElementSelected()
        {
            const string code = @"
Private Const myConst As Long = 42

Private myModuleVariable As Long


Public Sub DoIt()
    Dim myLocalVariable As Long
      
    myModuleVariable = myLocalVariable
End Sub";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(9, 5))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Procedure, "DoIt");

            Assert.AreEqual(expected, actual);
        }

        [Category("Resolver")]
        [Test]
        public void SelectionInsideVariableDeclaringReDimButNotOnIdentifier_ContainingModuleBodyElementSelected()
        {
            const string code = @"
Private Const myConst As Long = 42

Private myModuleVariable As Long


Public Sub DoIt()
    ReDim arr(23 To 42) As Long
      
    myModuleVariable = arr(33)
End Sub";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(8, 7))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Procedure, "DoIt");

            Assert.AreEqual(expected, actual);
        }

        [Category("Resolver")]
        [Test]
        public void SelectionInsideVariableDeclarationStatementForMultipleLocalVariables_ContainingModuleBodyElementSelected()
        {
            const string code = @"
Private Const myConst As Long = 42

Private myModuleVariable As Long


Public Sub DoIt()
    Dim myLocalVariable As Long, myOtherLocalVariable As String
      
    myModuleVariable = myLocalVariable
End Sub";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(8, 6))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Procedure, "DoIt");

            Assert.AreEqual(expected, actual);
        }

        [Category("Resolver")]
        [Test]
        public void SelectionInsideModuleBodyElementAroundVariableDeclarationButNotContainedInIt_ContainingModuleBodyElementSelected()
        {
            const string code = @"
Private Const myConst As Long = 42

Private myModuleVariable As Long


Public Sub DoIt()
    Dim myLocalVariable As Long
      
    myModuleVariable = myLocalVariable
End Sub";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(8, 1, 8, 32))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.Procedure, "DoIt");

            Assert.AreEqual(expected, actual);
        }

        [Category("Resolver")]
        [Test]
        public void SelectionOutsideModuleBodyElementAndOnNothingElse_ModuleSelected()
        {
            const string code = @"
Private Const myConst As Long = 42

Private myModuleVariable As Long


Public Sub DoIt()
    Dim myLocalVariable As Long
      
    myModuleVariable = myLocalVariable
End Sub";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(6, 1))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.ProceduralModule, "TestModule");

            Assert.AreEqual(expected, actual);
        }

        [Category("Resolver")]
        [Test]
        public void SelectionInsideVariableDeclarationStatementForMultipleModuleVariables_ModuleSelected()
        {
            const string code = @"
Private Const myConst As Long = 42

Private myModuleVariable As Long, myOtherModuleVariable As String


Public Sub DoIt()
    Dim myLocalVariable As Long
      
    myModuleVariable = myLocalVariable
End Sub";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(4, 2))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.ProceduralModule, "TestModule");

            Assert.AreEqual(expected, actual);
        }

        [Category("Resolver")]
        [Test]
        public void SelectionInsideConstantDeclarationStatementForMultipleModuleConstants_ModuleSelected()
        {
            const string code = @"
Private Const myConst As Long = 42, myOtherConstant As Long = 23

Private myModuleVariable As Long


Public Sub DoIt()
    Dim myLocalVariable As Long
      
    myModuleVariable = myLocalVariable
End Sub";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(2, 2))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.ProceduralModule, "TestModule");

            Assert.AreEqual(expected, actual);
        }

        [Category("Resolver")]
        [Test]
        public void SelectionAroundMemberButNotContained_ModuleSelected()
        {
            const string code = @"
Private Const myConst As Long = 42

Private myModuleVariable As Long


Public Sub DoIt()
    Dim myLocalVariable As Long
      
    myModuleVariable = myLocalVariable
End Sub";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(6, 1, 11,8))
                .AddProjectToVbeBuilder()
                .Build();

            var (expected, actual) = DeclarationsFromParse(vbe.Object, DeclarationType.ProceduralModule, "TestModule");

            Assert.AreEqual(expected, actual);
        }

        private static (Declaration specifiedDeclaration, Declaration selectedDeclaration) DeclarationsFromParse(
            IVBE vbe,
            DeclarationType declarationType, 
            string declarationName,
            string parentScope = null)
        {
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var selectionProvider = new SelectionService(vbe, state.ProjectsProvider);
                var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionProvider, state);

                var specifiedDeclaration = parentScope == null
                    ? state.DeclarationFinder
                        .DeclarationsWithType(declarationType)
                        .Single(declaration => declaration.IdentifierName.Equals(declarationName))
                    : state.DeclarationFinder
                        .DeclarationsWithType(declarationType)
                        .Single(declaration => declaration.IdentifierName.Equals(declarationName) 
                                               && declaration.ParentScope.Equals(parentScope));
                var selectedDeclaration = selectedDeclarationProvider.SelectedDeclaration();

                return (specifiedDeclaration, selectedDeclaration);
            }
        }
    }
}