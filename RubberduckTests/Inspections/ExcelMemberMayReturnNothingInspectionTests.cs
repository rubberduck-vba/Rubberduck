using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ExcelMemberMayReturnNothingInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsResult_FindWithMemberAccess()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    foo = ws.UsedRange.Find(""foo"").Row
End Sub
";
            Assert.AreEqual(1, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsResult_WithMemberAccessOnFind()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    With ws.UsedRange
        foo = .Find(""foo"").Row
    End With
End Sub
";
            Assert.AreEqual(1, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    '@Ignore ExcelMemberMayReturnNothing
    foo = ws.UsedRange.Find(""foo"").Row
End Sub
";

            Assert.AreEqual(0, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsNoResult_ResultIsNothingInAssignment()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    Dim foo As Boolean
    foo = ws.UsedRange.Find(""foo"") Is Nothing
End Sub
";

            Assert.AreEqual(0, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsNoResult_TransientAccessIsNothing()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    If ws.UsedRange.Find(""foo"") Is Nothing Then
        Debug.Print ""Not found""
    End If
End Sub
";

            Assert.AreEqual(0, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsNoResult_AssignedToVariableIsNothing()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    Dim result As Range
    Set result = ws.UsedRange.Find(""foo"")
    If result Is Nothing Then
        Debug.Print ""Not found""
    End If
End Sub
";

            Assert.AreEqual(0, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsResult_AssignedAndNotTested()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    Dim result As Range
    Set result = ws.UsedRange.Find(""foo"")
    result.Value = ""bar""
End Sub
";

            Assert.AreEqual(1, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsResult_AssignedAndNotTested_FromWithMemberAccess()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    Dim result As Range
    With ws.UsedRange
        Set result = .Find(""foo"")
    End With
    result.Value = ""bar""
End Sub
";

            Assert.AreEqual(1, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsResult_ResultIsSomethingElse()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    Dim result As Range
    Set result = ws.Range(""A1"")
    If ws.UsedRange.Find(""foo"") Is result Then
        Debug.Print ""Found it the dumb way""
    End If
End Sub
";

            Assert.AreEqual(1, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsResult_FindNext()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    Dim foo As Range
    Set foo = ws.UsedRange.Find(""foo"")
    If Not foo Is Nothing Then
        bar = ws.UsedRange.FindNext(foo).Row
    End If
End Sub
";

            Assert.AreEqual(1, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsResult_FindPrevious()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    Dim foo As Range
    Set foo = ws.UsedRange.Find(""foo"")
    If Not foo Is Nothing Then
        bar = ws.UsedRange.FindPrevious(foo).Row
    End If
End Sub
";

            Assert.AreEqual(1, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsResult_DefaultAccessExpression_Let()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    If ws.Range(""B:B"").Find(""bar"") = 1 Then 
        Debug.Print ""bar""
    End If
End Sub
";

            Assert.AreEqual(1, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsResult_DefaultAccessExpression_Argument()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    Bar ws.Range(""B:B"").Find(""bar"")
End Sub

Private Sub Bar(arg As Long)
End Sub
";

            Assert.AreEqual(1, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsResult_ObjectArgument()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    Bar ws.Range(""B:B"").Find(""bar"")
End Sub

Private Sub Bar(arg As Range)
End Sub
";

            Assert.AreEqual(1, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsResult_DefaultAccessExpression_Argument_ExplicitCall()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    Call Bar(ws.Range(""B:B"").Find(""bar""))
End Sub

Private Sub Bar(arg As Long)
End Sub
";

            Assert.AreEqual(1, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsResult_ObjectArgument_ExplicitCall()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    Call Bar(ws.Range(""B:B"").Find(""bar""))
End Sub

Private Sub Bar(arg As Range)
End Sub
";

            Assert.AreEqual(1, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsResult_WithBlock()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    With ws.Range(""B:B"").Find(""bar"")
    End With
End Sub

Private Sub Bar(arg As Range)
End Sub
";

            Assert.AreEqual(1, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_DoesNotReturnResult_CallStatement()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    ws.Range(""B:B"").Find ""bar""
End Sub

Private Sub Bar(arg As Range)
End Sub
";

            Assert.AreEqual(0, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_DoesNotReturnResult_ExplicitCallStatement()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    Call ws.Range(""B:B"").Find(""bar"")
End Sub

Private Sub Bar(arg As Range)
End Sub
";

            Assert.AreEqual(0, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsResult_FindAsWithBlockVariable()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    With ws.UsedRange.Find(""foo"")
        foo = .Row
    End With
End Sub
";

            Assert.AreEqual(1, InspectionResults(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsResult_AssignedToWithBlockVariable()
        {
            const string inputCode =
                @"Sub UnderTest()
    Dim ws As Worksheet
    Set ws = Sheet1
    Dim result As Range
    Set result = ws.UsedRange.Find(""foo"")
    With result
        .Value = ""bar""
    End With
End Sub
";

            Assert.AreEqual(1, InspectionResults(inputCode).Count());
        }

        private IEnumerable<IInspectionResult> InspectionResults(string inputCode)
            => InspectionResultsForModules(("Module1", inputCode, ComponentType.StandardModule), ReferenceLibrary.Excel);

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ExcelMemberMayReturnNothingInspection(state);
        }
    }
}
