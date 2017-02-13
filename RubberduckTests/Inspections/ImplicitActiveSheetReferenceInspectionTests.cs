using System.Diagnostics;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ImplicitActiveSheetReferenceInspectionTests
    {
        [TestMethod]
        [DeploymentItem(@"TestFiles\")]
        [TestCategory("Inspections")]
        public void ImplicitActiveSheetReference_ReportsRange()
        {
            const string inputCode =
@"Sub foo()
    Dim arr1() As Variant
    arr1 = Range(""A1:B2"")
End Sub
";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)                
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));
            parser.State.AddTestLibrary("Excel.1.8.xml");

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitActiveSheetReferenceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [DeploymentItem(@"TestFiles\")]
        [TestCategory("Inspections")]
        public void ImplicitActiveSheetReference_Ignored_DoesNotReportRange()
        {
            const string inputCode =
@"Sub foo()
    Dim arr1() As Variant

    '@Ignore ImplicitActiveSheetReference
    arr1 = Range(""A1:B2"")
End Sub
";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitActiveSheetReferenceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [DeploymentItem(@"TestFiles\")]
        public void ImplicitActiveSheetReference_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Sub foo()
    Dim arr1() As Variant
    arr1 = Range(""A1:B2"")
End Sub";

            const string expectedCode =
@"Sub foo()
    Dim arr1() As Variant
'@Ignore ImplicitActiveSheetReference
    arr1 = Range(""A1:B2"")
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();
            var module = project.Object.VBComponents[0].CodeModule;
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));
            parser.State.AddTestLibrary("Excel.1.8.xml");

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitActiveSheetReferenceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.Single(s => s is IgnoreOnceQuickFix).Fix();

            Assert.AreEqual(expectedCode, module.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new ImplicitActiveSheetReferenceInspection(null);
            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ImplicitActiveSheetReferenceInspection";
            var inspection = new ImplicitActiveSheetReferenceInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}