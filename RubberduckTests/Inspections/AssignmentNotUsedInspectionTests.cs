using System.Collections.Generic;
using NUnit.Framework;
using Rubberduck.Inspections.CodePathAnalysis;
using Rubberduck.Inspections.Concrete;
using RubberduckTests.Mocks;
using System.Linq;
using System.Threading;
using Rubberduck.Parsing.Inspections.Abstract;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    [Category("Inspections")]
    [Category("AssignmentNotUsed")]
    public class AssignmentNotUsedInspectionTests
    {
        private IEnumerable<IInspectionResult> GetInspectionResults(string code, bool includeLibraries = false)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _, referenceStdLibs: includeLibraries);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new AssignmentNotUsedInspection(state, new ProcedureTreeVisitor());
                var inspector = InspectionsHelper.GetInspector(inspection);
                return inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            }
        }

        [Test]
        public void IgnoresExplicitArrays()
        {
            const string code = @"
Sub Foo()
    Dim bar(1 To 10) As String
    bar(1) = 42
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void IgnoresImplicitArrays()
        {
            const string code = @"
Sub Foo()
    Dim bar As Variant
    ReDim bar(1 To 10)
    bar(1) = ""Z""
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void IgnoresImplicitReDimmedArray()
        {
            const string code = @"
Sub Test()
    Dim foo As Variant
    ReDim foo(1 To 10)
    foo(1) = 42
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void MarksSequentialAssignments()
        {
            const string code = @"
Sub Foo()
    Dim i As Integer
    i = 9
    i = 8
    Bar i
End Sub

Sub Bar(ByVal i As Integer)
End Sub";
            var results = GetInspectionResults(code);
            Assert.AreEqual(1, results.Count());
        }

        [Test]
        public void MarksLastAssignmentInDeclarationBlock()
        {
            const string code = @"
Sub Foo()
    Dim i As Integer
    i = 9
End Sub";
            var results = GetInspectionResults(code);
            Assert.AreEqual(1, results.Count());
        }

        [Test]
        public void DoesNotMarkUsedAssignmentForVariableDeclaredAndAssignedInsideInnerBlock()
        {
            const string code = @"
Public Function IssueInIfBlock(ByVal someInput As Boolean) As Boolean
    
    If someInput Then
        Dim HasProblems As Boolean
        HasProblems = True ' triggers ""Assignment is not used"" inspection
    End If
    
    IssueInIfBlock = someInput And HasProblems
    
End Function";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void DoesNotMarkSuccessiveReferencingAssignments()
        {
            const string code = @"
Public Sub Foo()
    Dim bar As String
    bar = ""a""
    bar = bar & ""b""
    bar = bar & ""c""
    Debug.Print bar
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void MarksUnusedAssignmentsInIfElseBlock()
        {
            const string code = @"
Sub Foo()
    Dim i As Integer
    i = 0
    If i = 9 Then '<~ reads i=0
        i = 8 '<~ unused
    Else
        i = 1 '<~ unused
    End If
End Sub";
            var results = GetInspectionResults(code);
            Assert.AreEqual(2, results.Count());
        }

        [Test]
        public void DoesNotMarkAssignmentWithReferenceAfter()
        {
            const string code = @"
Sub Foo()
    Dim i As Integer
    i = 9
    Bar i
End Sub

Sub Bar(ByVal i As Integer)
End Sub";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void DoesNotMarkAssignment_UsedInForNext()
        {
            const string code = @"
Sub foo()
    Dim i As Integer
    i = 1
    For counter = i To 2
    Next
End Sub";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void MarksUnusedAssignment_InWhileWend()
        {
            const string code = @"
Sub foo()
    Dim i As Integer, j As Integer
    i = 0

    While i < 10 '<~ i=0 has a read here
        i = i + 1 '<~ current i has a read here
        j = i '<~ not used
    Wend
End Sub";
            var results = GetInspectionResults(code);
            Assert.AreEqual(1, results.Count());
        }

        [Test]
        public void DoesNotMarkAssignment_UsedInDoWhile()
        {
            const string code = @"
Sub foo()
    Dim i As Integer
    i = 0
    Do While i < 10 ' <~ i=0 has a read here
    Loop
End Sub";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void DoesNotMarkAssignment_UsedInSelectCase()
        {
            const string code = @"
Sub foo()
    Dim i As Integer
    i = 0
    Select Case i ' <~ i=0 has a read here
        Case 0
            i = 1
        Case 1
            i = 2
        Case 2
            i = 3
        Case Else
            i = -1
    End Select
End Sub";
            var results = GetInspectionResults(code);
            Assert.AreEqual(4, results.Count());
        }

        [Test]
        public void DoesNotMarkAssignment_UsingNothing()
        {
            const string code = @"
Public Sub Test()
Dim my_fso As Scripting.FileSystemObject
Set my_fso = New Scripting.FileSystemObject

Debug.Print my_fso.GetFolder(""C:\Windows"").DateLastModified

Set my_fso = Nothing
End Sub";
            var results = GetInspectionResults(code, includeLibraries:true);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void DoesMarkAssignment_UsingNothing_NotUsed()
        {
            const string code = @"
Public Sub Test()
Dim my_fso As Scripting.FileSystemObject
Set my_fso = New Scripting.FileSystemObject

Set my_fso = Nothing
End Sub";
            var results = GetInspectionResults(code, includeLibraries: true);
            Assert.AreEqual(1, results.Count());
        }

        [Test]
        public void MarksUnusedAssignmentInDoLoopBody()
        {
            const string code = @"
Public Sub IssueInDoLoop()

    Dim Idx As Long
    Idx = 1
    Do
        Dim SomeBreakCondition As Boolean
        SomeBreakCondition = True '<~ should not trip inspection...
        Idx = Idx + 1 '<~ not used
    Loop Until SomeBreakCondition '<~ ...because of this use
    
End Sub";
            var results = GetInspectionResults(code);
            Assert.AreEqual(1, results.Count());
        }

        [Test]
        public void DoesNotMarkResults_InIgnoredModule()
        {
            const string code = @"'@IgnoreModule 
Public Sub Test()
    Dim foo As Long
    foo = 1245316
End Sub";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void DoesNotMarkAssignment_WithIgnoreOnceAnnotation()
        {
            const string code = @"Public Sub Test()
    Dim foo As Long
    '@Ignore AssignmentNotUsed
    foo = 123451
    foo = 56126 '<~ not used
End Sub";
            var results = GetInspectionResults(code);
            Assert.AreEqual(1, results.Count());
        }

        [Test]
        public void DoesNotMarkAssignment_ToIgnoredDeclaration()
        {
            const string code = @"Public Sub Test()
    '@Ignore AssignmentNotUsed
    Dim foo As Long
    foo = 123467
    foo = 45678
End Sub";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void DoesNotMarkAssignment_ToStaticVariable()
        {
            const string code = @"Public Function IssueWithStaticVariable() As Boolean

    Static SomeProperty As String
    If SomeProperty = ""dummy"" Then
        IssueWithStaticVariable = True
    End If
    SomeProperty = ""new dynamic value"" ' triggers ""Assignment is not used"" inspection
    
End Function";

            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        // todo handle static variables
        //[Ignore("Static variables are ignored for now.")]
        public void /*Marks*/IgnoresUnusedStaticVariableAssignment()
        {
            const string code = @"Public Sub Test()
Static foo As Long
foo = 42 ' <~ blatantly not used
foo = 0 ' <~ not used either
End Sub";
            var results = GetInspectionResults(code);
            Assert.AreEqual(/*2*/ 0, results.Count());
        }

        [Test]
        public void MarksElseBlockAssignmentNotUsed()
        {
            const string code = @"Public Sub Test()
Dim foo As Long
foo = 42 '<~ will be considered as not used
If True Then
  Debug.Print foo
Else
  foo = 10 '<~ not used
End If
End Sub";
            var results = GetInspectionResults(code);
            Assert.AreEqual(2, results.Count());
        }

        [Test]
        public void MarksUnusedConditionalVariableAssignment()
        {
            const string code = @"Public Sub Test()
Dim foo As Long
If True Then ' <~ condition expression is not evaluated
foo = 42 ' <~ blatantly not used
foo = 0 ' <~ not used either
End If
End Sub";
            var results = GetInspectionResults(code);
            Assert.AreEqual(2, results.Count());
        }

        [Test]
        public void MarksUnusedNestedConditionalVariableAssignment()
        {
            const string code = @"Public Sub Test()
Dim foo As Long
If foo > 0 Or foo <> foo + 1 Then ' <~ condition expression is not evaluated, but that's 3 refs against no assignment
  If False Then
    foo = 42 ' <~ not used
  End If
  foo = 0
  Debug.Print foo
End If
End Sub";
            var results = GetInspectionResults(code);
            Assert.AreEqual(1, results.Count());
        }
    }
}
