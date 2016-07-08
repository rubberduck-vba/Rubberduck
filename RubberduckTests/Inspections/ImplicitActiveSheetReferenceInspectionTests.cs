using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ImplicitActiveSheetReferenceInspectionTests
    {
        [TestMethod]    // doesn't pick up the reference to "Range".
        [TestCategory("Inspections")]
        public void ReportsRange()
        {
            const string inputCode =
@"Sub foo()
    Dim arr1() As Variant
    arr1 = Range(""A1:B2"")
End Sub
";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .AddReference("Excel", "C:\\Program Files\\Microsoft Office\\Root\\Office 16\\EXCEL.EXE", true)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object, new Mock<ISinks>().Object));

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

            parser.State.AddDeclaration(excelDeclaration);
            parser.State.AddDeclaration(globalDeclaration);
            parser.State.AddDeclaration(globalCoClassDeclaration);
            parser.State.AddDeclaration(rangeClassModuleDeclaration);
            parser.State.AddDeclaration(rangeDeclaration);
            parser.State.AddDeclaration(firstParamDeclaration);
            parser.State.AddDeclaration(secondParamDeclaration);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitActiveSheetReferenceInspection(vbe.Object, parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }
    }
}