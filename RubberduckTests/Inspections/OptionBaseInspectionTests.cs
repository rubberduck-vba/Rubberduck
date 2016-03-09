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
    public class OptionBaseInspectionTests
    {
        [TestMethod]
        public void OptionBaseOneSpecified_ReturnsResult()
        {
            const string inputCode = @"Option Base 1";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.ParseSynchronous();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new OptionBaseInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void OptionBaseNotSpecified_DoesNotReturnResult()
        {
            const string inputCode = @"";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.ParseSynchronous();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new OptionBaseInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void OptionBaseZeroSpecified_DoesNotReturnResult()
        {
            const string inputCode = @"Option Base 0";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.ParseSynchronous();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new OptionBaseInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void OptionBaseOneSpecified_ReturnsMultipleResults()
        {
            const string inputCode = @"Option Base 1";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.ParseSynchronous();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new OptionBaseInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        public void OptionBaseOnePartiallySpecified_ReturnsResults()
        {
            const string inputCode1 = @"";
            const string inputCode2 = @"Option Base 1";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.ParseSynchronous();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new OptionBaseInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void InspectionType()
        {
            var inspection = new OptionBaseInspection(null);
            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
        }

        [TestMethod]
        public void InspectionName()
        {
            const string inspectionName = "OptionBaseInspection";
            var inspection = new OptionBaseInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}