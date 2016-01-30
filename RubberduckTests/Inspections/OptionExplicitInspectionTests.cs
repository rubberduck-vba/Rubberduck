using System.Linq;
using System.Threading;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class OptionExplicitInspectionTests
    {
        private readonly SemaphoreSlim _semaphore = new SemaphoreSlim(0, 1);

        void State_StateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State == ParserState.Ready)
            {
                _semaphore.Release();
            }
        }

        [TestMethod, Timeout(1000)]
        public void NotAlreadySpecified_ReturnsResult()
        {
            const string inputCode = @"";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var inspection = new OptionExplicitInspection(parseResult.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod, Timeout(1000)]
        public void AlreadySpecified_DoesNotReturnResult()
        {
            const string inputCode = @"Option Explicit";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var inspection = new OptionExplicitInspection(parseResult.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod, Timeout(1000)]
        public void NotAlreadySpecified_ReturnsMultipleResults()
        {
            const string inputCode = @"";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var inspection = new OptionExplicitInspection(parseResult.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod, Timeout(1000)]
        public void PartiallySpecified_ReturnsResults()
        {
            const string inputCode1 = @"";
            const string inputCode2 = @"Option Explicit";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var inspection = new OptionExplicitInspection(parseResult.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod, Timeout(1000)]
        public void NotAlreadySpecified_QuickFixWorks()
        {
            const string inputCode = @"";
            const string expectedCode =
@"Option Explicit

";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var inspection = new OptionExplicitInspection(parseResult.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void InspectionType()
        {
            var inspection = new OptionExplicitInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod, Timeout(1000)]
        public void InspectionName()
        {
            const string inspectionName = "OptionExplicitInspection";
            var inspection = new OptionExplicitInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}