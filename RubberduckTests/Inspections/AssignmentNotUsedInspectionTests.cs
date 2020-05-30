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
        public void IgnoresAssignmentsWithinBranchNodes()
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
            Assert.AreEqual(0, results.Count(), TestFailureMsg(results, ""));
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
        [TestCase("Do While", "Loop")]
        public void IgnoresAssignmentsWithinLoopNodes(string loopBeginStmt, string loopEndStmt)
        {
            var code = $@"
Sub foo()
    Dim i As Integer
    i = 0
    {loopBeginStmt}
        i = i + 1
    {loopEndStmt}
End Sub";
            var results = InspectionResultsForStandardModule(code);
            Assert.AreEqual(0, results.Count(), TestFailureMsg(results, ""));
        }

        [Test]
        public void DoesNotMarkAssignment_UsedInSelectCase()
        {
            const string code = @"
Sub foo()
    Dim i As Integer
    Dim unused As String
    unused = ""Unused""
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

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new AssignmentNotUsedInspection(state, new Walker());
        }
    }
}
