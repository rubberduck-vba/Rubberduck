using System.Collections.Generic;
using NUnit.Framework;
using Rubberduck.Inspections.CodePathAnalysis;
using RubberduckTests.Mocks;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    [Category("Inspections")]
    [Category("AssignmentNotUsed")]
    public class AssignmentNotUsedInspectionTests : InspectionTestsBase
    {
        private IEnumerable<IInspectionResult> InspectionResultsForStandardModuleUsingStdLibs(string code)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _, referenceStdLibs: true).Object;
            return InspectionResults(vbe);
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
            var results = InspectionResultsForStandardModule(code);
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
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void IgnoresImplicitReDimmedArray()
        {
            const string code = @"
Sub test()
    Dim foo As Variant
    ReDim foo(1 To 10)
    foo(1) = 42
End Sub
";
            var results = InspectionResultsForStandardModule(code);
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
            var results = InspectionResultsForStandardModule(code);
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
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(1, results.Count());
        }

        [Test]
        // Note: both assignments in the if/else can be marked in the future;
        // I just want feedback before I start mucking around that deep.
        public void DoesNotMarkLastAssignmentInNonDeclarationBlock()
        {
            const string code = @"
Sub Foo()
    Dim i As Integer
    i = 0
    If i = 9 Then
        i = 8
    Else
        i = 8
    End If
End Sub";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(0, results.Count());
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
            var results = InspectionResultsForStandardModule(code);
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
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void DoesNotMarkAssignment_UsedInWhileWend()
        {
            const string code = @"
Sub foo()
    Dim i As Integer
    i = 0

    While i < 10
        i = i + 1
    Wend
End Sub";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void DoesNotMarkAssignment_UsedInDoWhile()
        {
            const string code = @"
Sub foo()
    Dim i As Integer
    i = 0
    Do While i < 10
    Loop
End Sub";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void DoesNotMarkAssignment_UsedInSelectCase()
        {
            const string code = @"
Sub foo()
    Dim i As Integer
    i = 0
    Select Case i
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
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void DoesNotMarkAssignment_UsingNothing()
        {
            const string code = @"
Public Sub Test()
Dim my_fso   As Scripting.FileSystemObject
Set my_fso = New Scripting.FileSystemObject

Debug.Print my_fso.GetFolder(""C:\Windows"").DateLastModified

Set my_fso = Nothing
End Sub";
            var results = InspectionResultsForStandardModuleUsingStdLibs(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void DoesMarkAssignment_UsingNothing_NotUsed()
        {
            const string code = @"
Public Sub Test()
Dim my_fso   As Scripting.FileSystemObject
Set my_fso = New Scripting.FileSystemObject

Set my_fso = Nothing
End Sub";
            var results = InspectionResultsForStandardModuleUsingStdLibs(code);
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
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void DoesNotMarkAssignment_WithIgnoreOnceAnnotation()
        {
            const string code = @"Public Sub Test()
    Dim foo As Long
    '@Ignore AssignmentNotUsed
    foo = 123451
    foo = 56126
End Sub";
            var results = InspectionResultsForStandardModule(code);
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
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(0, results.Count());
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/5456
        [TestCase("Resume CleanExit")]
        [TestCase("GoTo CleanExit")]
        [TestCase("Resume 8")] //Inverse = ratio
        [TestCase("GoTo 8")] //Inverse = ratio
        public void IgnoresAssignmentWhereUsedByJumpStatement(string statement)
        {
            string code =
$@"
Public Function Inverse(value As Double) As Double
    Dim ratio As Double
    ratio = 0# 'assigment not used - flagged
On Error Goto ErrorHandler
    ratio = 1# / value
CleanExit:
    Inverse = ratio
    Exit Function
ErrorHandler:
    'assigment not used evaluation disqualified by Resume/GoTo - not flagged
    ratio = -1#
    {statement}
End Function
";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(1, results.Count());
        }

        [TestCase("Resume CleanExit")]
        [TestCase("GoTo CleanExit")]
        [TestCase("Resume 8")] //Inverse = ratio
        [TestCase("GoTo 8")] //Inverse = ratio
        public void IgnoresAssignmentWhereUsedByJumpStatement_JumpOnSameLineAsAssignment(string statement)
        {
            string code =
$@"
Public Function Inverse(value As Double) As Double
    Dim ratio As Double
    ratio = 0# 'assigment not used - flagged
On Error Goto ErrorHandler
    ratio = 1# / value
CleanExit:
    Inverse = ratio
    Exit Function
ErrorHandler:
    'assigment not used evaluation disqualified by Resume/GoTo - not flagged
    ratio = -1#: {statement}
End Function
";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(1, results.Count());
        }

        [TestCase("GoTo")]
        [TestCase("Resume")]
        public void MultipleSingleLineJumpStmts(string jumpStatement)
        {
            string code =
$@"
Public Function Fizz(value As Double) As Double
    Dim firstVal As Double
    Dim anotherVal As Double
    Dim yetAnotherVal As Double
On Error GoTo ErrorHandler
    Fizz = 1# / value
    Exit Function
Exit1:
    Fizz = anotherVal
    Exit Function 
Exit2:
    Fizz = firstVal
    Exit Function
Exit3:
    Fizz = yetAnotherVal
    Exit Function
ErrorHandler:
    anotherVal = 6#: {jumpStatement} Exit1: firstVal = -1#: {jumpStatement} Exit2: yetAnotherVal = -99#: {jumpStatement} Exit3
End Function
";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(0, results.Count());
        }


        [TestCase("Exit Function: Fizz = firstVal: Exit Function", 1)] //value not read
        [TestCase("Fizz = firstVal: Exit Function", 0)] //value is read
        public void ExitStmtOnSameLineAsNonAssignments(string exitFunctionLine, int expected)
        {
            string code =
$@"
Public Function Fizz(value As Double) As Double
    Dim firstVal As Double
On Error GoTo ErrorHandler
    Fizz = 1# / value
    Exit Function
Exit1:
    Fizz = 0#
    {exitFunctionLine} 
ErrorHandler:
    firstVal = -1#: GoTo Exit1
End Function
";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(expected, results.Count());
        }

        [Test]
        public void IgnoresExitForStmt()
        {
            string code =
$@"
Public Function DoStuff() As Double
On Error Goto ErrorHandler
    Dim fizz As Double
    Dim index As Long
    For index = 0 to 20
        fizz = CDbl(index)
        If index = 5 Then
            Exit For
        End If
    Next
Finally:
    DoStuff = fizz
    Exit Function
ErrorHandler:
    fizz = 100#
    Resume Next
End Function
";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void IgnoresExitDoStmt()
        {
            string code =
$@"
Public Function DoStuff() As Double
On Error Goto ErrorHandler
    Dim fizz As Double
    Dim index As Long
    index = 0
    Do While index < 20
        fizz = CDbl(index)
        If index = 5 Then
            Exit Do
        End If
        index = index + 1
    Loop 
Finally:
    DoStuff = fizz
    Exit Function
ErrorHandler:
    fizz = 100#
    Resume Next
End Function
";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void GotoRespectsExitStmt()
        {
            string code =
$@"
Public Function DoStuff(value As Double) As Double
    Dim fizz As Double
    GoTo DumbGotoLabel2

DumbGotoLabel1:
    DoStuff = 6#
    Exit Function

    DoStuff = fizz
    Exit Function

DumbGotoLabel2:
    fizz = value
    GoTo DumbGotoLabel1
End Function
";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(1, results.Count());
        }

        [Test]
        [TestCase("Resume IgnoreRatio")]
        [TestCase("GoTo IgnoreRatio")]
        public void FlagsAssignmentWhereExecutionPathModifiedByJumpStatementCouldNotIncludeUse(string statement)
        {
            string code =
$@"
Public Function Inverse(value As Double) As Double
    Inverse = 0#
    Dim ratio As Double
On Error Goto ErrorHandler
    ratio = 1# / value
    Inverse = ratio

IgnoreRatio:
    Exit Function
ErrorHandler:
    'assignment not used since jump is to IgnoreRatio: 
    'and all ratio references below IgnoreRatio: are assignments
    ratio = 0# 
    {statement}
End Function
";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(1, results.Count());
        }

        [Test]
        public void FlagsWhereNoJumpStatementsFollowsUnusedAssignment()
        {
            string code =
$@"
Public Function Inverse(value As Double) As Double
    Inverse = 0#
    Dim ratio As Double
On Error Goto ErrorHandler
    ratio = 1# / value
IgnoreRatio:
    Inverse = ratio
    Exit Function
ErrorHandler:
    ratio = 0# 'Removed from unused results
    Resume IgnoreRatio
    ratio = 0# 'Not Used
End Function
";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(1, results.Count());
        }

        [Test]
        [TestCase("Resume Next")]
        [TestCase("Resume")]
        public void ResumeStmtSpecialCases(string resumeStmt)
        {
            string code =
$@"
Public Function Inverse(value As Double) As Double
    Inverse = 0#
    Dim ratio As Double
On Error Goto ErrorHandler
    ratio = 1# / value
    Inverse = ratio
    Exit Function
ErrorHandler:
    ratio = 0#
    {resumeStmt}
End Function
";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [TestCase("Resume Next")]
        [TestCase("Resume")]
        public void ResumeStmt_VariableReadIsNotAvailableInExecutionBranch(string resumeStmt)
        {
            string code =
$@"
Public Function Inverse(value As Double) As Double
    Inverse = 0#
    Dim ratio As Double
On Error GoTo 0
    ratio = 1# / value
    Inverse = ratio
On Error GoTo ErrorHandler:
    Exit Function
ErrorHandler:
    ratio = 0#
    {resumeStmt}
End Function
";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(1, results.Count());
        }


        [Test]
        public void ResumeStmt_OnSameLineAsLabel()
        {
            string code =
$@"
Public Function Inverse(value As Double) As Double
    Dim ratio As Double
On Error GoTo ErrorHandler:
    ratio = 1# / value
    Inverse = ratio
    Exit Function
ErrorHandler: ratio = 0#: Resume Next
End Function
";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [TestCase("Resume Next", true, 2)]
        [TestCase("Resume", true, 2)]
        [TestCase("Resume Next", false, 1)]
        [TestCase("Resume", false, 1)]
        public void ResumeStmt_NarrowsEvaluationsUsingExitStatements(string resumeStmt, bool errorHandler2HasExit, int expected)
        {
            string exitStmt = errorHandler2HasExit ? "Exit Function" : "'Exit Function";
            string code =
$@"
Public Function Inverse(value As Double) As Double
    Dim ratio As Double
    Inverse = 0#
On Error GoTo ErrorHandler1:
    ratio = 1# / value
    Inverse = ratio
On Error GoTo ErrorHandler2:
    ratio = 0# 'Not used
    {exitStmt}
On Error GoTo ErrorHandler3:
    ratio = 0# * 34# 'Used
    Inverse = ratio
    Exit Function
ErrorHandler1:
    ratio = 0# 'Used
    {resumeStmt}
ErrorHandler2:
    ratio = 0# 'Not used with ""Exit Function"", removed from unused without ""Exit Function""
    {resumeStmt}
ErrorHandler3:
End Function
";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(expected, results.Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new AssignmentNotUsedInspection(state, new Walker());
        }
    }
}
