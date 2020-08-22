using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.MoveCloserToUsage;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class MoveCloserToUsageTests : RefactoringTestBase
    {
        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_ModuleVariable_ClearsResidualNewLines()
        {
            //Input
            const string inputCode =
@"Private bar As Boolean





Private Sub Foo()
    bar = True
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar As Boolean
    bar = True
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_LocalVariable_ClearsResidualNewLines()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
    Dim bar As Boolean

    Dim var1 As Long

    Dim var2 As String

    bar = True
End Sub";
            var selection = new Selection(2, 10);

            //Expectation
            const string expectedCode =
@"Private Sub Foo()
    Dim var1 As Long

    Dim var2 As String

    Dim bar As Boolean
    bar = True
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_Field()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean


Private Sub Foo()
    bar = True
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar As Boolean
    bar = True
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_LineNumbers()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo()
100 bar = True
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar As Boolean
100 bar = True
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_Field_MultipleLines()
        {
            //Input
            const string inputCode =
                @"Private _
bar _
As _
Boolean
Private Sub Foo()
    bar = True
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar As Boolean
    bar = True
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_FieldInOtherClass()
        {
            //Input
            const string inputModuleCode =
                @"Public bar As Boolean";

            const string inputClassCode =
                @"Private Sub Foo()
Module1.bar = True
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedModuleCode =
                @"";

            const string expectedClassCode =
                @"Private Sub Foo()
Dim bar As Boolean
bar = True
End Sub";

            var actualCode = RefactoredCode(
                "Module1",
                selection,
                null,
                false,
                ("Module1", inputModuleCode, ComponentType.StandardModule),
                ("Class1", inputClassCode, ComponentType.ClassModule));
            Assert.AreEqual(expectedModuleCode, actualCode["Module1"]);
            Assert.AreEqual(expectedClassCode, actualCode["Class1"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_Variable()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
    Dim bar As Boolean
    Dim bat As Integer
    bar = True
End Sub";
            var selection = new Selection(4, 6, 4, 8);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bat As Integer
    Dim bar As Boolean
    bar = True
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_VariableWithLineNumbers()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
1   Dim bar As Boolean
2   Dim bat As Integer
3   bar = True
End Sub";
            var selection = new Selection(4, 6, 4, 8);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
1   
2   Dim bat As Integer
    Dim bar As Boolean
3   bar = True
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_Variable_MultipleLines()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
    Dim _
    bar _
    As _
    Boolean
    Dim bat As Integer
    bar = True
End Sub";
            var selection = new Selection(4, 6, 4, 8);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bat As Integer
    Dim bar As Boolean
    bar = True
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_MultipleFields_MoveSecond()
        {
            //Input
            const string inputCode =
                @"Private bar As Integer
Private bat As Boolean
Private bay As Date

Private Sub Foo()
    bat = True
End Sub";
            var selection = new Selection(2, 1);

            //Expectation
            const string expectedCode =
                @"Private bar As Integer
Private bay As Date

Private Sub Foo()
    Dim bat As Boolean
    bat = True
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_MultipleFieldsOneStatement_MoveFirst()
        {
            //Input
            const string inputCode =
                @"Private bar As Integer, _
          bat As Boolean, _
          bay As Date

Private Sub Foo()
    bar = 3
End Sub";
            var selection = new Selection(6, 6);

            //Expectation
            const string expectedCode =
                @"Private bat As Boolean, _
          bay As Date

Private Sub Foo()
    Dim bar As Integer
    bar = 3
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_MultipleFieldsOneStatement_MoveSecond()
        {
            //Input
            const string inputCode =
                @"Private bar As Integer, _
          bat As Boolean, _
          bay As Date

Private Sub Foo()
    bat = True
End Sub";
            var selection = new Selection(6, 6);

            //Expectation
            const string expectedCode =
                @"Private bar As Integer, _
          bay As Date

Private Sub Foo()
    Dim bat As Boolean
    bat = True
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_MultipleFieldsOneStatement_MoveLast()
        {
            //Input
            const string inputCode =
                @"Private bar As Integer, _
          bat As Boolean, _
          bay As Date

Private Sub Foo()
    bay = #1/13/2004#
End Sub";
            var selection = new Selection(6, 6);

            //Expectation
            const string expectedCode =
                @"Private bar As Integer, _
          bat As Boolean

Private Sub Foo()
    Dim bay As Date
    bay = #1/13/2004#
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_MultipleVariablesOneStatement_MoveFirst()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
    Dim bar As Integer, _
        bat As Boolean, _
        bay As Date

    bat = True
    bar = 3
End Sub";
            var selection = new Selection(2, 9);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bat As Boolean, _
        bay As Date

    bat = True
    Dim bar As Integer
    bar = 3
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_MultipleVariablesOneStatement_MoveSecond()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
    Dim bar As Integer, _
        bat As Boolean, _
        bay As Date

    bar = 1
    bat = True
End Sub";
            var selection = new Selection(3, 10);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar As Integer, _
        bay As Date

    bar = 1
    Dim bat As Boolean
    bat = True
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_MultipleVariablesOneStatement_MoveLast()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
    Dim bar As Integer, _
        bat As Boolean, _
        bay As Date

    bar = 4
    bay = #1/13/2004#
End Sub";
            var selection = new Selection(4, 11);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar As Integer, _
        bat As Boolean

    bar = 4
    Dim bay As Date
    bay = #1/13/2004#
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_ReferenceIsNotBeginningOfStatement_Assignment()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo(ByRef bat As Boolean)
    bat = bar
End Sub";

            const string expectedCode =
                @"Private Sub Foo(ByRef bat As Boolean)
    Dim bar As Boolean
    bat = bar
End Sub";
            var selection = new Selection(1, 1);

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_ReferenceIsNotBeginningOfStatement_PassAsParam()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo()
    Baz bar
End Sub
Sub Baz(ByVal bat As Boolean)
End Sub";

            const string expectedCode =
                @"Private Sub Foo()
    Dim bar As Boolean
    Baz bar
End Sub
Sub Baz(ByVal bat As Boolean)
End Sub";
            var selection = new Selection(1, 1);

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_ReferenceIsNotBeginningOfStatement_PassAsParam_ReferenceIsNotFirstLine()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo()
    Baz True, _
        True, _
        bar
End Sub
Sub Baz(ByVal bat As Boolean, ByVal bas As Boolean, ByVal bac As Boolean)
End Sub";

            const string expectedCode =
                @"Private Sub Foo()
    Dim bar As Boolean
    Baz True, _
        True, _
        bar
End Sub
Sub Baz(ByVal bat As Boolean, ByVal bas As Boolean, ByVal bac As Boolean)
End Sub";
            var selection = new Selection(1, 1);

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_ReferenceIsSeparatedWithColon()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo(): Baz True, True, bar: End Sub
Private Sub Baz(ByVal bat As Boolean, ByVal bas As Boolean, ByVal bac As Boolean): End Sub";

            var selection = new Selection(1, 1);

            // Yeah, this code is a mess.  That is why we got the SmartIndenter
            const string expectedCode =
                @"Private Sub Foo(): Dim bar As Boolean : Baz True, True, bar: End Sub
Private Sub Baz(ByVal bat As Boolean, ByVal bas As Boolean, ByVal bac As Boolean): End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_WorksWithNamedParameters()
        {
            //Input
            const string inputCode =
                @"
Private foo As Long

Public Sub Test()
    SomeSub someParam:=foo
End Sub

Public Sub SomeSub(ByVal someParam As Long)
    Debug.Print someParam
End Sub";

            var selection = new Selection(2, 1);
            const string expectedCode =
                @"
Public Sub Test()
    Dim foo As Long
    SomeSub someParam:=foo
End Sub

Public Sub SomeSub(ByVal someParam As Long)
    Debug.Print someParam
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_WorksWithNamedParametersAndStatementSeparators()
        {
            //Input
            const string inputCode =
                @"Private foo As Long

Public Sub Test(): SomeSub someParam:=foo: End Sub

Public Sub SomeSub(ByVal someParam As Long)
    Debug.Print someParam
End Sub";

            var selection = new Selection(1, 1);
            const string expectedCode =

@"Public Sub Test(): Dim foo As Long : SomeSub someParam:=foo: End Sub

Public Sub SomeSub(ByVal someParam As Long)
    Debug.Print someParam
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void IntroduceFieldRefactoring_PassInTarget_NonVariable()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo()
    bar = True
End Sub";

            var actualCode = RefactoredCode(inputCode, "Foo", DeclarationType.Procedure, typeof(InvalidDeclarationTypeException));
            Assert.AreEqual(inputCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void IntroduceFieldRefactoring_DeclarationOfInvalidTypeSelected()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo()
    bar = True
End Sub";
            var selection = new Selection(2, 15);

            var actualCode = RefactoredCode(inputCode, selection, typeof(NoDeclarationForSelectionException));
            Assert.AreEqual(inputCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_NoReferences()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo()
End Sub";
            var selection = new Selection(1, 1);

            var actualCode = RefactoredCode(inputCode, selection, typeof(TargetDeclarationNotUsedException));
            Assert.AreEqual(inputCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_ReferencedInMultipleProcedures()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo()
    bar = True
End Sub
Private Sub Bar()
    bar = True
End Sub";
            var selection = new Selection(1, 1);

            var actualCode = RefactoredCode(inputCode, selection, typeof(TargetDeclarationUsedInMultipleMethodsException));
            Assert.AreEqual(inputCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_VariableWithSameNameAlreadyExistsInProcedure()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
    Dim bar As Boolean
    OtherModule.Bar = True
End Sub";
            var selection = new Selection(3, 18);

            const string otherModuleInputCode =
                @"Public Bar As Boolean";

            var actualCode = RefactoredCode(
                "Module", 
                selection, 
                typeof(TargetDeclarationConflictsWithPreexistingDeclaration),
                false,
                ("Module", inputCode, ComponentType.StandardModule),
                ("OtherModule", otherModuleInputCode, ComponentType.StandardModule));
            Assert.AreEqual(inputCode, actualCode["Module"]);
            Assert.AreEqual(otherModuleInputCode, actualCode["OtherModule"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_ModuleVariableWithSameNameAlreadyExists()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo()
    OtherModule.Bar = True
End Sub";
            var selection = new Selection(3, 18);

            const string otherModuleInputCode =
                @"Public Bar As Boolean";

            var actualCode = RefactoredCode(
                "Module",
                selection,
                typeof(TargetDeclarationConflictsWithPreexistingDeclaration),
                false,
                ("Module", inputCode, ComponentType.StandardModule),
                ("OtherModule", otherModuleInputCode, ComponentType.StandardModule));
            Assert.AreEqual(inputCode, actualCode["Module"]);
            Assert.AreEqual(otherModuleInputCode, actualCode["OtherModule"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_NonPrivateInNonStandardModule()
        {
            //Input
            const string inputCode =
                @"Public bar As Boolean
Private Sub Foo()
    bar = True
End Sub";
            var selection = new Selection(1, 9);

            var actualCode = RefactoredCode(
                "Class",
                selection,
                typeof(TargetDeclarationNonPrivateInNonStandardModule),
                false,
                ("Class", inputCode, ComponentType.ClassModule));
            Assert.AreEqual(inputCode, actualCode["Class"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_TargetInDifferentNonStandardModule()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
    Dim baz As Class1
    baz.Bar = True
End Sub";
            var selection = new Selection(3, 10);

            const string otherClassInputCode =
                @"Public Bar As Boolean";

            var actualCode = RefactoredCode(
                "Module",
                selection,
                typeof(TargetDeclarationInDifferentNonStandardModuleException),
                false,
                ("Module", inputCode, ComponentType.StandardModule),
                ("Class1", otherClassInputCode, ComponentType.ClassModule));
            Assert.AreEqual(inputCode, actualCode["Module"]);
            Assert.AreEqual(otherClassInputCode, actualCode["Class1"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_TargetInDifferentProject()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
    OtherProject.OtherModule.Bar = True
End Sub";
            var selection = new Selection(2, 31);

            const string otherModuleInputCode =
                @"Public Bar As Boolean";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("OtherProject", "otherProjectPath",ProjectProtection.Unprotected)
                .AddComponent("OtherModule", ComponentType.StandardModule, otherModuleInputCode)
                .AddProjectToVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Module", ComponentType.StandardModule, inputCode)
                .AddReference("OtherProject", "otherProjectPath", 0,0,false,ReferenceKind.Project)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var module = state.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule).Single(decl => decl.IdentifierName == "Module").QualifiedModuleName;
                var qualifiedSelection = new QualifiedSelection(module, selection);
                var testRefactoring = TestRefactoring(rewritingManager, state);

                Assert.Throws<TargetDeclarationInDifferentProjectThanUses>(() =>
                    testRefactoring.Refactor(qualifiedSelection));
                
                var otherModule = state.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule).Single(decl => decl.IdentifierName == "OtherModule").QualifiedModuleName;
                var actualModuleCode = state.ProjectsProvider.Component(module).CodeModule.Content();
                var actualOtherModuleCode = state.ProjectsProvider.Component(otherModule).CodeModule.Content();

                Assert.AreEqual(inputCode, actualModuleCode);
                Assert.AreEqual(otherModuleInputCode, actualOtherModuleCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_TargetNotUserDefined()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
    Dim baz As Excel.Range
    baz.Value = 42
End Sub";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Module", ComponentType.StandardModule, inputCode)
                .AddReference(ReferenceLibrary.Excel)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var target = state.DeclarationFinder.MatchName("Value")
                    .First(decl => decl.ParentDeclaration.IdentifierName == "Range");
                var testRefactoring = TestRefactoring(rewritingManager, state);

                Assert.Throws<TargetDeclarationNotUserDefinedException>(() =>
                    testRefactoring.Refactor(target));

                var module = state.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule).Single(decl => decl.IdentifierName == "Module").QualifiedModuleName;
                var actualModuleCode = state.ProjectsProvider.Component(module).CodeModule.Content();

                Assert.AreEqual(inputCode, actualModuleCode);
            }
        }

        [Test]
        [Category("Move Closer")]
        [Category("Refactorings")]
        public void MoveCloser_RespectsMemberAccess_ContextOwners()
        {
            const string inputCode =
@"
Public Sub foo()
  Dim count As Long
  Dim report As Worksheet
  Set report = ThisWorkbook.ActiveSheet
  With report
    For count = 1 To 10
      If .Cells(1, count) > count Then
        .Cells(2, count).Value2 = count
      End If
    Next
  End With
End Sub";
            var selection = new Selection(3, 8);

            const string expectedCode =
@"
Public Sub foo()
  Dim report As Worksheet
  Set report = ThisWorkbook.ActiveSheet
  With report
    Dim count As Long
    For count = 1 To 10
      If .Cells(1, count) > count Then
        .Cells(2, count).Value2 = count
      End If
    Next
  End With
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("Move Closer")]
        [Category("Refactorings")]
        public void MoveCloser_RespectsObjectProperties_InUsages()
        {
            string inputClassCode =
@"
Option Explicit

Private _name As Long
Private _myOtherProperty As Long

Public Property Set Name(name As String)
    _name = name
End Property

Public Property Get Name() As String
    Name = _name
End Property

Public Property Set OtherProperty(val As Long)
    _myOtherProperty = val
End Property

Public Property Get OtherProperty() As Long
    OtherProperty = _myOtherProperty
End Property

";
            string inputCode = @"Private foo As Class1


Public Sub Test()
    Debug.Print ""Some statements between""
    Debug.Print ""Declaration and first usage!""
    Set foo = new Class1
    foo.Name = ""FooName""
    foo.OtherProperty = 1626
End Sub";

            var selection = new Selection(1, 1);

            const string expectedCode = @"Public Sub Test()
    Debug.Print ""Some statements between""
    Debug.Print ""Declaration and first usage!""
    Dim foo As Class1
    Set foo = new Class1
    foo.Name = ""FooName""
    foo.OtherProperty = 1626
End Sub";

            var actualCode = RefactoredCode(
                "Module1", 
                selection, 
                null, 
                false, 
                ("Module1", inputCode, ComponentType.StandardModule), 
                ("Class1", inputClassCode, ComponentType.ClassModule));
            Assert.AreEqual(expectedCode, actualCode["Module1"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_DynamicArray()
        {
            //Input
            const string inputCode =
                @"Private bar() As Boolean
Private Sub Foo()
    ReDim bar(0)
    bar(0) = True
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar() As Boolean
    ReDim bar(0)
    bar(0) = True
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_FixedArray()
        {
            //Input
            const string inputCode =
                @"Private bar(0) As Boolean
Private Sub Foo()
    bar(0) = True
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar(0) As Boolean
    bar(0) = True
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_FixedArrayBounded()
        {
            //Input
            const string inputCode =
                @"Private bar(1 To 42) As Boolean
Private Sub Foo()
    bar(1) = True
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar(1 To 42) As Boolean
    bar(1) = True
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_MultiDimensionalArray()
        {
            //Input
            const string inputCode =
                @"Private bar(1, 1) As Boolean
Private Sub Foo()
    bar(0, 0) = True
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar(1, 1) As Boolean
    bar(0, 0) = True
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_SelfAssigned()
        {
            //Input
            const string inputCode =
                @"Private bar As New Collection
Private Sub Foo()
    bar.Add 42
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar As New Collection
    bar.Add 42
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, ISelectionService selectionService)
        {
            var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionService, state);
            var baseRefactoring = new MoveCloserToUsageRefactoringAction(rewritingManager);
            return new MoveCloserToUsageRefactoring(baseRefactoring, state, selectionService, selectedDeclarationProvider);
        }
    }
}
