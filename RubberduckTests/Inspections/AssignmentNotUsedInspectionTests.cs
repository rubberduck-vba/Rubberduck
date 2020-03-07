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

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new AssignmentNotUsedInspection(state, new Walker());
        }
    }
}
