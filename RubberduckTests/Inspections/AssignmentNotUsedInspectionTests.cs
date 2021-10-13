using System.Collections.Generic;
using NUnit.Framework;
using Rubberduck.Inspections.CodePathAnalysis;
using RubberduckTests.Mocks;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

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
            Assert.AreEqual(0, results.Count(), TestFailureMsg(results, ""));
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
            Assert.AreEqual(0, results.Count(), TestFailureMsg(results, ""));
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
            Assert.AreEqual(0, results.Count(), TestFailureMsg(results, ""));
        }

        [Test]
        public void MarksSequentialAssignments()
        {
            var expectedResults = new string[] { "i", "unused" };
            var code = @"
Sub Foo()
    Dim unused As String
    unused = ""Unused""
    Dim i As Integer
    i = 9
    i = 8
    Bar i
End Sub

Sub Bar(ByVal i As Integer)
End Sub";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(expectedResults.Count(), results.Count(), TestFailureMsg(results, expectedResults));
        }

        [Test]
        public void IgnoresAssignmentsWithinBranchNodes()
        {
            const string code = @"
Sub Foo()
    Dim unused As String
    unused = ""Unused""
    Dim i As Integer
    i = 0
    If i = 9 Then
        i = 8
    Else
        i = 8
    End If
End Sub";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(1, results.Count(), TestFailureMsg(results, "unused"));
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
            Assert.AreEqual(0, results.Count(), TestFailureMsg(results, ""));
        }

        [TestCase("For counter = 1 To 20", "Next")]
        [TestCase("While i < 10", "Wend")]
        [TestCase("Do While i < 10", "Loop")]
        public void IgnoresAssignmentsWithinLoopNodes(string loopBeginStmt, string loopEndStmt)
        {
            var code = $@"
Sub foo()
    Dim i As Integer
    i = 0
    {loopBeginStmt}
        i = i + 1
    {loopEndStmt}
    Dim unused As String
    unused = ""Unused""
End Sub";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(1, results.Count(), TestFailureMsg(results, "unused"));
        }

        [Test]
        public void DoesNotMarkAssignment_UsedInSelectCase()
        {
            const string code = @"
Sub foo()
    Dim unused As String
    unused = ""Unused""
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
            Assert.AreEqual(1, results.Count(), TestFailureMsg(results, "unused"));
        }

        [Test]
        public void DoesNotMarkAssignment_UsingNothing()
        {
            const string code = @"
Public Sub Test()
Dim my_fso   As Scripting.FileSystemObject
Dim unused As String
unused = ""Unused""
Set my_fso = New Scripting.FileSystemObject

Debug.Print my_fso.GetFolder(""C:\Windows"").DateLastModified

Set my_fso = Nothing
End Sub";
            var results = InspectionResultsForStandardModuleUsingStdLibs(code);
            Assert.AreEqual(1, results.Count(), TestFailureMsg(results, "unused"));
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

        [TestCase("'@IgnoreModule", 0)]
        [TestCase("", 1)]
        public void DoesNotMarkResults_InIgnoredModule(string annotation, int expected)
        {
            var code = 
$@"{annotation} 
Public Sub Test()
    Dim foo As Long
    foo = 1245316
End Sub";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(expected, results.Count());
        }

        [TestCase("'@Ignore AssignmentNotUsed", 0)]
        [TestCase("", 1)]
        public void DoesNotMarkAssignment_WithIgnoreOnceAnnotation(string annotation, int expectedCount)
        {
            var code =
$@"Public Function Test() As Long
    Dim foo As Long
    {annotation}
    foo = 123451
    foo = 56126
    Test = foo
End Function";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(expectedCount, results.Count());
        }

        [TestCase("'@Ignore AssignmentNotUsed", "unused")]
        [TestCase("", "unused", "foo", "foo")]
        public void DoesNotMarkAssignment_ToIgnoredDeclaration(string annotation, params string[] expectedResults)
        {
            var code = $@"
Public Sub Test()
    {annotation}
    Dim foo As Long
    Dim unused As String
    unused = ""Unused""
    foo = 123467
    foo = 45678
End Sub";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(expectedResults.Count(), results.Count(), TestFailureMsg(results, expectedResults));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/4913
        [TestCase("Test = foo", "unused")]
        [TestCase("", "unused", "foo")]
        public void AssignmentUsedByNextLetAssignment(string finalStatement, params string[] expectedResults)
        {
            var code =
$@"
Public Function Test() As String
    Test = ""test""
    Dim foo As String
    Dim unused As String
    unused = ""Unused""
    foo = ""bar""
    foo = foo & ""baz""
    {finalStatement}
End Function";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(expectedResults.Count(), results.Count(), TestFailureMsg(results, expectedResults));
        }

        [TestCase("Test = foo", "foo", "foo", "unused")]
        [TestCase("", "foo", "foo", "foo", "unused")]
        public void AssignmentUsedByNextLetAssignment_HasSubsequentUnused(string finalStatement, params string[] expectedResults)
        {
            var code =
$@"
Public Function Test() As String
    Dim foo As String
    Dim unused As String
    unused = ""Unused""
    foo = ""bar"" 
    foo = foo & ""baz"" 'not used
    foo = ""Yo"" 'not used
    foo = ""YoYo"" 'depends
    {finalStatement}
End Function";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(expectedResults.Count(), results.Count(), TestFailureMsg(results, expectedResults));
        }

        [TestCase("temp = testValue + value", "unused")]
        [TestCase("temp = value", "unused", "testValue")]
        public void StaticVariable(string reference, params string[] expectedResults)
        {
            var code =
$@"
Public Function Test(value As Long) As Long
    Static testValue As Long
    Dim temp As Long
    {reference}
    Dim unused As String
    unused = ""Unused""
    testValue = temp 
    Test = temp
End Function";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(expectedResults.Count(), results.Count(), TestFailureMsg(results, expectedResults));
        }


        [TestCase("temp = testValue + value", "unused")]
        [TestCase("temp = value", "unused", "testValue")]
        public void StaticProcedure(string reference, params string[] expectedResults)
        {
            var code =
$@"
Public Static Function Test(value As Long) As Long
    Dim testValue As Long
    Dim temp As Long
    {reference}
    Dim unused As String
    unused = ""Unused""
    testValue = temp 
    Test = temp
End Function";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(expectedResults.Count(), results.Count(), TestFailureMsg(results, expectedResults));
        }

        [TestCase("Test = foo.ReturnOne()", "unused")]
        [TestCase(@"Test = ""test""", "unused", "foo")]
        public void AssignmentUsedByNextSetAssignment(string finalStatement, params string[] expectedResults)
        {
            var code =
$@"
Public Function Test() As String
    Dim unused As String
    unused = ""Unused""
    Dim foo As Class1
    Set foo = new Class1
    Set foo = foo
    {finalStatement}
End Function";

            var classModuleName = "Class1";
            var classCode =
$@"
Option Explicit
    
Public Function ReturnOne() As String
    ReturnOne = ""1""
End Function
";

            var results = InspectionResultsForModules(
                            (MockVbeBuilder.TestModuleName, code, ComponentType.StandardModule),
                            (classModuleName, classCode, ComponentType.ClassModule));

            Assert.AreEqual(expectedResults.Count(), results.Count(), TestFailureMsg(results, expectedResults));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/4913
        [Test]
        public void LocalVariablesHaveProcedureScope_WithBlock()
        {
            var code =
$@"
Public Function Test() As String
    Dim unused As String
    unused = ""Unused""  
    With Module2
        Dim localName As String
        localName = .Name
    End With
    Test = localName
End Function";

            var otherModuleName = "Module2";
            var otherModuleCode =
$@"
Option Explicit
    
Public Name As String
";

            var results = InspectionResultsForModules(
                            (MockVbeBuilder.TestModuleName, code, ComponentType.StandardModule),
                            (otherModuleName, otherModuleCode, ComponentType.StandardModule));

            Assert.AreEqual(1, results.Count(), TestFailureMsg(results, "unused"));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/4913
        [Test]
        public void LocalVariablesHaveProcedureScope_BranchNode()
        {
            var code =
@"
Public Sub Test(markerLocation As Long, componentName As String)
    Dim unused As String
    unused = ""Unused""
    If markerLocation > 0 Then
        Dim workingName As String
        workingName = Right$(componentName, Len(componentName) - markerLocation - 1)
    Else
        workingName = componentName
    End If
    markerLocation = InStr(1, workingName, ""."")
End Sub";

            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(1, results.Count(), TestFailureMsg(results, "unused"));
        }

        private static string TestFailureMsg(IEnumerable<IInspectionResult> results, params string[] expectedResults)
        {
            var resultStrings = results.Select(r => r.Target.IdentifierName);
            var expectedMissing = string.Join(", ", expectedResults.Except(expectedResults.Intersect(resultStrings)));
            if (!string.IsNullOrEmpty(expectedMissing))
            {
                return $"No result for: {expectedMissing}";
            }

            var unexpectedResults = string.Join(", ", resultStrings.Except(resultStrings.Intersect(expectedResults)));
            if (!string.IsNullOrEmpty(unexpectedResults))
            {
                return $"Unexpected result: {unexpectedResults}";
            }
            return string.Empty;
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/5456
        [TestCase("Resume CleanExit")]
        [TestCase("GoTo CleanExit")]
        [TestCase("Resume 850")]
        [TestCase("GoTo 850")]
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
850:
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
        [TestCase("Resume 850")]
        [TestCase("GoTo 850")]
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
850:
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
        [TestCase("Resume Next", 2)]
        [TestCase("Resume", 2)]
        public void ResumeStmt_NarrowsEvaluationsUsingExitStatementPrecedesUse(string resumeStmt, int expected)
        {
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
    Exit Function
On Error GoTo ErrorHandler3:
    ratio = 0# * 34# 'Used
    Inverse = ratio
    Exit Function
ErrorHandler1:
    ratio = 0# 'Used
    {resumeStmt}
ErrorHandler2:
    ratio = 0# 'Not used
    {resumeStmt}
ErrorHandler3:
End Function
";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(expected, results.Count());
        }

        [Test]
        [TestCase("Resume Next", 1)]
        [TestCase("Resume", 1)]
        public void ResumeStmt_NarrowsEvaluationsUsingExitStatements_FallsThroughToReference(string resumeStmt, int expected)
        {
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
On Error GoTo ErrorHandler3:
    ratio = 0# * 34#
    Inverse = ratio
    Exit Function
ErrorHandler1:
    ratio = 0#
    {resumeStmt}
ErrorHandler2:
    ratio = 0#
    {resumeStmt}
ErrorHandler3:
End Function
";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(expected, results.Count());
        }

        [Test]
        public void ConditionalAssignments_NotConsideredForOverwritingAssignment()
        {
            string code =
$@"
Public Function Test() As Boolean
    Dim value As Boolean
    value = True

    If True Then
        value = False
    End If

    Test = value
End Function
";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        public void ResumeStmt_SingleResult()
        {
            string code =
$@"
Public Function Inverse(value As Double) As Double
On Error GoTo ErrorHandler:
    Dim ratio As Double
    ratio = 0# '<== unused
    ratio = 1# / value '<== used
    Inverse = ratio
    Exit Function
ErrorHandler: 
    ratio = 0#  '<== possibly used
    Resume
End Function
";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(1, results.Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new AssignmentNotUsedInspection(state, new Walker());
        }
    }
}
