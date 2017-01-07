using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
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
                .AddReference("Excel", "C:\\Program Files\\Microsoft Office\\Root\\Office 16\\EXCEL.EXE", true)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            GetExcelRangeDeclarations().ForEach(d => parser.State.AddDeclaration(d));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitActiveSheetReferenceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
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
                .AddReference("Excel", "C:\\Program Files\\Microsoft Office\\Root\\Office 16\\EXCEL.EXE", true)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            GetExcelRangeDeclarations().ForEach(d => parser.State.AddDeclaration(d));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitActiveSheetReferenceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
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
                .AddReference("Excel", "C:\\Program Files\\Microsoft Office\\Root\\Office 16\\EXCEL.EXE", true)
                .Build();
            var module = project.Object.VBComponents[0].CodeModule;
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            GetExcelRangeDeclarations().ForEach(d => parser.State.AddDeclaration(d));

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
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, string.Empty)
                .AddReference("Excel", "C:\\Program Files\\Microsoft Office\\Root\\Office 16\\EXCEL.EXE", true)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var inspection = new ImplicitActiveSheetReferenceInspection(null);
            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, string.Empty)
                .AddReference("Excel", "C:\\Program Files\\Microsoft Office\\Root\\Office 16\\EXCEL.EXE", true)
                .Build();
            var vbe = builder.AddProject(project).Build();

            const string inspectionName = "ImplicitActiveSheetReferenceInspection";
            var inspection = new ImplicitActiveSheetReferenceInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        private List<Declaration> GetExcelRangeDeclarations()
        {
            var excelDeclaration = new ProjectDeclaration(new QualifiedMemberName(new QualifiedModuleName("Excel",
                    "C:\\Program Files\\Microsoft Office\\Root\\Office 16\\EXCEL.EXE", "Excel"), "Excel"), "Excel", true);

            var globalDeclaration = new ClassModuleDeclaration(new QualifiedMemberName(
                new QualifiedModuleName("Excel",
                    "C:\\Program Files\\Microsoft Office\\Root\\Office 16\\EXCEL.EXE", "_Global"),
                "_Global"), excelDeclaration, "_Global", true, null, null);

            var globalCoClassDeclarationAttributes = new Attributes();
            globalCoClassDeclarationAttributes.AddPredeclaredIdTypeAttribute();
            globalCoClassDeclarationAttributes.AddGlobalClassAttribute();

            var globalCoClassDeclaration = new ClassModuleDeclaration(new QualifiedMemberName(
                new QualifiedModuleName("Excel",
                    "C:\\Program Files\\Microsoft Office\\Root\\Office 16\\EXCEL.EXE", "Global"),
                "Global"), excelDeclaration, "Global", true, null, globalCoClassDeclarationAttributes);

            globalDeclaration.AddSubtype(globalCoClassDeclaration);
            globalCoClassDeclaration.AddSupertype(globalDeclaration);
            globalCoClassDeclaration.AddSupertype("_Global");

            var rangeClassModuleDeclaration = new ClassModuleDeclaration(new QualifiedMemberName(
                new QualifiedModuleName("Excel",
                    "C:\\Program Files\\Microsoft Office\\Root\\Office 16\\EXCEL.EXE", "Range"),
                "Range"), excelDeclaration, "Range", true, new List<IAnnotation>(), new Attributes());

            var rangeDeclaration = new PropertyGetDeclaration(new QualifiedMemberName(
                new QualifiedModuleName("Excel",
                    "C:\\Program Files\\Microsoft Office\\Root\\Office 16\\EXCEL.EXE", "_Global"), "Range"),
                globalDeclaration, globalDeclaration, "Range", null, null, Accessibility.Global, null, Selection.Home,
                false, true, new List<IAnnotation>(), new Attributes());

            var firstParamDeclaration = new ParameterDeclaration(new QualifiedMemberName(
                new QualifiedModuleName("Excel",
                    "C:\\Program Files\\Microsoft Office\\Root\\Office 16\\EXCEL.EXE", "_Global"),
                "Cell1"), rangeDeclaration, "Variant", null, null, false, false);

            var secondParamDeclaration = new ParameterDeclaration(new QualifiedMemberName(
                new QualifiedModuleName("Excel",
                    "C:\\Program Files\\Microsoft Office\\Root\\Office 16\\EXCEL.EXE", "_Global"),
                "Cell2"), rangeDeclaration, "Variant", null, null, true, false);

            rangeDeclaration.AddParameter(firstParamDeclaration);
            rangeDeclaration.AddParameter(secondParamDeclaration);

            return new List<Declaration>
            {
                excelDeclaration,
                globalDeclaration,
                globalCoClassDeclaration,
                rangeClassModuleDeclaration,
                rangeDeclaration,
                firstParamDeclaration,
                secondParamDeclaration,
            };
        }
    }
}