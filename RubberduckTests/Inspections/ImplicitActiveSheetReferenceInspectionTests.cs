using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
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
        [TestMethod, Ignore]    // doesn't pick up the reference to "Range".
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
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .AddReference("Excel", string.Empty, true)

                // Apparently, the COM loader can't find it when it isn't actually loaded...
                //.AddReference("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", true)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            var excelDeclaration = new ProjectDeclaration(new QualifiedMemberName(new QualifiedModuleName("Excel",
                    "C:\\Program Files\\Microsoft Office\\Root\\Office 16\\EXCEL.EXE", "Excel"), "Excel"), "Excel", true);

            var listColumnDeclaration = new ClassModuleDeclaration(new QualifiedMemberName(
                new QualifiedModuleName("Excel",
                    "C:\\Program Files\\Microsoft Office\\Root\\Office 16\\EXCEL.EXE", "ListColumn"),
                "ListColumn"), excelDeclaration, "ListColumn", true, null, null, true, true);

            var rangeDeclaration =
                new Declaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("Excel",
                            "C:\\Program Files\\Microsoft Office\\Root\\Office 16\\EXCEL.EXE", "ListColumn"), "Range"),
                    listColumnDeclaration, "EXCEL.EXE;Excel.ListColumn", "Range", false, false, Accessibility.Global,
                    (DeclarationType)3712, true, null, new Attributes());

            parser.State.AddDeclaration(excelDeclaration);
            parser.State.AddDeclaration(listColumnDeclaration);
            parser.State.AddDeclaration(rangeDeclaration);

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitActiveSheetReferenceInspection(vbe.Object, parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }
    }
}