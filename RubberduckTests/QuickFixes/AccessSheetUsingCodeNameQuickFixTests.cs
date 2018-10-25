using System.Linq;
using System.Threading;
using Moq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class AccessSheetUsingCodeNameQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void SheetAccessedUsingString_QuickFixWorks_UsingSheetThroughWorkbookModule()
        {
            const string inputCode = @"
Public Sub Foo()
    ThisWorkbook.Sheets(""Sheet1"").Range(""A1"") = ""foo""
End Sub";

            const string expectedCode = @"
Public Sub Foo()
    Sheet1.Range(""A1"") = ""foo""
End Sub";

            using (var state = ArrangeParserAndParse(inputCode, out var component))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new AccessSheetUsingCodeNameQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void SheetAccessedUsingString_QuickFixWorks_UsingSheetThroughApplicationModule()
        {
            const string inputCode = @"
Public Sub Foo()
    Application.Sheets(""Sheet1"").Range(""A1"") = ""foo""
End Sub";

            const string expectedCode = @"
Public Sub Foo()
    Sheet1.Range(""A1"") = ""foo""
End Sub";

            using (var state = ArrangeParserAndParse(inputCode, out var component))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new AccessSheetUsingCodeNameQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void SheetAccessedUsingString_QuickFixWorks_UsingSheetThroughGlobalModule()
        {
            const string inputCode = @"
Public Sub Foo()
    Sheets(""Sheet1"").Range(""A1"") = ""foo""
End Sub";

            const string expectedCode = @"
Public Sub Foo()
    Sheet1.Range(""A1"") = ""foo""
End Sub";

            using (var state = ArrangeParserAndParse(inputCode, out var component))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new AccessSheetUsingCodeNameQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void SheetAccessedUsingString_QuickFixWorks_AssigningSheetToVariable()
        {
            const string inputCode = @"
Public Sub Foo()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(""Sheet1"")
    ws.Cells(1, 1) = ""foo""
    Bar ws
End Sub

Public Sub Bar(ws As Worksheet)
End Sub";

            const string expectedCode = @"
Public Sub Foo()
    
    
    Sheet1.Cells(1, 1) = ""foo""
    Bar Sheet1
End Sub

Public Sub Bar(ws As Worksheet)
End Sub";
            
            using (var state = ArrangeParserAndParse(inputCode, out var component))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new AccessSheetUsingCodeNameQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void SheetAccessedUsingString_QuickFixWorks_AssigningSheetToUnusedVariable()
        {
            const string inputCode = @"
Public Sub Foo()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(""Sheet1"")
End Sub";

            const string expectedCode = @"
Public Sub Foo()
    
    
End Sub";

            using (var state = ArrangeParserAndParse(inputCode, out var component))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new AccessSheetUsingCodeNameQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void SheetAccessedUsingString_QuickFixWorks_AssigningSheetToVariableDeclaredInList()
        {
            const string inputCode = @"
Public Sub Foo()
    Dim ws As Worksheet, s As String
    Set ws = ThisWorkbook.Worksheets(""Sheet1"")
    ws.Cells(1, 1) = ""foo""
End Sub";

            const string expectedCode = @"
Public Sub Foo()
    Dim s As String
    
    Sheet1.Cells(1, 1) = ""foo""
End Sub";

            using (var state = ArrangeParserAndParse(inputCode, out var component))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new AccessSheetUsingCodeNameQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void SheetAccessedUsingString_QuickFixWorks_AssigningSheetToVariableDeclaredLastInList()
        {
            const string inputCode = @"
Public Sub Foo()
    Dim s As String, ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(""Sheet1"")
    ws.Cells(1, 1) = ""foo""
End Sub";

            const string expectedCode = @"
Public Sub Foo()
    Dim s As String
    
    Sheet1.Cells(1, 1) = ""foo""
End Sub";

            using (var state = ArrangeParserAndParse(inputCode, out var component))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new AccessSheetUsingCodeNameQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void SheetAccessedUsingString_QuickFixWorks_SheetVariableWithSameNameAsOtherDeclarations()
        {
            const string inputCode = @"
Public Sub Foo()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(""Sheet1"")
    ws.Cells(1, 1) = ""foo""
End Sub

Public Sub ws()
    Dim ws As Worksheet
End Sub";

            const string expectedCode = @"
Public Sub Foo()
    
    
    Sheet1.Cells(1, 1) = ""foo""
End Sub

Public Sub ws()
    Dim ws As Worksheet
End Sub";

            using (var state = ArrangeParserAndParse(inputCode, out var component))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new AccessSheetUsingCodeNameQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void SheetAccessedUsingString_QuickFixWorks_SheetNameDifferentThanSheetCodeName()
        {
            const string inputCode = @"
Public Sub Foo()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(""Name"")
    ws.Cells(1, 1) = ""foo""
End Sub

Public Sub ws()
    Dim ws As Worksheet
End Sub";

            const string expectedCode = @"
Public Sub Foo()
    
    
    CodeName.Cells(1, 1) = ""foo""
End Sub

Public Sub ws()
    Dim ws As Worksheet
End Sub";

            using (var state = ArrangeParserAndParse(inputCode, out var component))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new AccessSheetUsingCodeNameQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void SheetAccessedUsingString_QuickFixWorks_TransientReferenceSetStatement()
        {
            const string inputCode = @"
Sub Test()
    Dim ws As Worksheet
    Set ws = Worksheets.Add(Worksheets(""Sheet1""))
    Debug.Print ws.Name
End Sub";

            const string expectedCode = @"
Sub Test()
    Dim ws As Worksheet
    Set ws = Worksheets.Add(Sheet1)
    Debug.Print ws.Name
End Sub";

            using (var state = ArrangeParserAndParse(inputCode, out var component))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new AccessSheetUsingCodeNameQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void SheetAccessedUsingString_QuickFixWorks_TransientReferenceNoSetStatement()
        {
            const string inputCode = @"
Sub Test()
    If Not Worksheets.Add(Worksheets(""Sheet1"")) Is Nothing Then
        Debug.Print ""Added""
    End If
End Sub";

            const string expectedCode = @"
Sub Test()
    If Not Worksheets.Add(Sheet1) Is Nothing Then
        Debug.Print ""Added""
    End If
End Sub";

            using (var state = ArrangeParserAndParse(inputCode, out var component))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new AccessSheetUsingCodeNameQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void SheetAccessedUsingString_QuickFixWorks_ImplicitVariableAssignment()
        {
            const string inputCode = @"
Sub Test()
    Set ws = Worksheets(""Sheet1"")
    ws.Name = ""Foo""
End Sub";

            const string expectedCode = @"
Sub Test()
    
    Sheet1.Name = ""Foo""
End Sub";

            using (var state = ArrangeParserAndParse(inputCode, out var component))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new AccessSheetUsingCodeNameQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        private static RubberduckParserState ArrangeParserAndParse(string inputCode, out IVBComponent component)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddComponent("Sheet1", ComponentType.Document, "",
                    properties: new[]
                    {
                        CreateVBComponentPropertyMock("Name", "Sheet1").Object,
                        CreateVBComponentPropertyMock("CodeName", "Sheet1").Object
                    })
                .AddComponent("SheetWithDifferentCodeName", ComponentType.Document, "",
                    properties: new[]
                    {
                        CreateVBComponentPropertyMock("Name", "Name").Object,
                        CreateVBComponentPropertyMock("CodeName", "CodeName").Object
                    })
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();

            component = project.Object.VBComponents[0];

            var vbe = builder.AddProject(project).Build();

            return MockParser.CreateAndParse(vbe.Object);
        }

        private static Mock<IProperty> CreateVBComponentPropertyMock(string propertyName, string propertyValue)
        {
            var propertyMock = new Mock<IProperty>();
            propertyMock.SetupGet(m => m.Name).Returns(propertyName);
            propertyMock.SetupGet(m => m.Value).Returns(propertyValue);

            return propertyMock;
        }
    }
}
