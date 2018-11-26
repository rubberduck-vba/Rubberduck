using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ExcelMemberMayReturnNothingInspectionTests
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

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ExcelMemberMayReturnNothingInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ExcelMemberMayReturnNothingInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ExcelMemberMayReturnNothingInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ExcelMemberMayReturnNothingInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ExcelMemberMayReturnNothingInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ExcelMemberMayReturnNothingInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ExcelMemberMayReturnNothingInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ExcelMemberMayReturnNothingInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ExcelMemberMayReturnNothingInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ExcelMemberMayReturnNothing_ReturnsResult_DefaultAccessExpression()
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

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ExcelMemberMayReturnNothingInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ExcelMemberMayReturnNothingInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ExcelMemberMayReturnNothingInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        private static RubberduckParserState ArrangeParserAndParse(string inputCode)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();

            var vbe = builder.AddProject(project).Build();

            return MockParser.CreateAndParse(vbe.Object); ;
        }
    }
}
