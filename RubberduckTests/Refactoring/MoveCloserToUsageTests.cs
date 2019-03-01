using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.MoveCloserToUsage;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class MoveCloserToUsageTests : RefactoringTestBase
    {
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
            var selection = new Selection(2, 16);

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
            var selection = new Selection(3, 16);

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
            var selection = new Selection(4, 16);

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
                @"
Public Sub Test(): Dim foo As Long : SomeSub someParam:=foo: End Sub

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
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {
                var refactoring = TestRefactoring(rewritingManager, state);
                Assert.Throws<InvalidDeclarationTypeException>(() => refactoring.Refactor(state.AllUserDeclarations.First(d => d.DeclarationType != DeclarationType.Variable)));

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(inputCode, actualCode);
            }
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

            const string expectedCode = @"

Public Sub Test()
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
            return new MoveCloserToUsageRefactoring(state, rewritingManager, selectionService);
        }
    }
}
