using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class UseMeaningfulNameInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void UseMeaningfulName_ReturnsResult_NameWithoutVowels()
        {
            const string inputCode = 
@"Sub Ffffff()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("MyClass", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UseMeaningfulNameInspection(null, parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UseMeaningfulName_ReturnsResult_NameUnderThreeLetters()
        {
            const string inputCode =
@"Sub Oo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("MyClass", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UseMeaningfulNameInspection(null, parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UseMeaningfulName_ReturnsResult_NameEndsWithDigit()
        {
            const string inputCode =
@"Sub Foo1()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("MyClass", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UseMeaningfulNameInspection(null, parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UseMeaningfulName_DoesNotReturnsResult_GoodName_LowerCaseVowels()
        {
            const string inputCode =
@"Sub FooBar()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("MyClass", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UseMeaningfulNameInspection(null, parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UseMeaningfulName_DoesNotReturnsResult_GoodName_UpperCaseVowels()
        {
            const string inputCode =
@"Sub FOOBAR()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("MyClass", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UseMeaningfulNameInspection(null, parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UseMeaningfulName_DoesNotReturnsResult_OptionBase()
        {
            const string inputCode =
@"Option Base 1";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("MyClass", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UseMeaningfulNameInspection(null, parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new UseMeaningfulNameInspection(null, null);
            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "UseMeaningfulNameInspection";
            var inspection = new UseMeaningfulNameInspection(null, null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
