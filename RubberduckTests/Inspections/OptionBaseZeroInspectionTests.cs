using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;
using Rubberduck.Settings;
using System.Threading;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete.Rubberduck.Inspections;
using Rubberduck.Inspections.Resources;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class OptionBaseZeroInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void OptionBaseZeroStatement_ReturnsResult()
        {
            const string inputCode =
@"Option Base 0";

            //Arrange
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new OptionBaseZeroInspection(parser.State);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(parser.State, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void OptionBaseOneStatement_NoResult()
        {
            string inputCode = "Option Base 1";

            //Arrange
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new OptionBaseZeroInspection(parser.State);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(parser.State, CancellationToken.None).Result;

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void OptionBaseZeroStatement_Ignored_DoesNotReturnResult()
        {
            string inputCode =
@"'@Ignore OptionBaseZero
Option Base 0";

            //Arrange
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteCallStatementInspection(parser.State);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(parser.State, CancellationToken.None).Result;

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void OptionBaseZeroStatement_QuickFixWorks_RemoveStatement()
        {
            string inputCode = "Option Base 0";
            string expectedCode = string.Empty;

            //Arrange
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new OptionBaseZeroInspection(parser.State);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(parser.State, CancellationToken.None).Result;

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.First().Fix();
            }

            var actual = module.Content();
            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void OptionBaseZeroStatement_QuickFixWorks_RemoveStatement_MultipleLines()
        {
            string inputCode =
@"Option _
Base _
0";

            string expectedCode = string.Empty;

            //Arrange
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new OptionBaseZeroInspection(parser.State);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(parser.State, CancellationToken.None).Result;

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.First().Fix();
            }

            var actual = module.Content();
            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void OptionBaseZeroStatement_QuickFixWorks_RemoveStatement_InstructionSeparator()
        {
            string inputCode = "Option Explicit: Option Base 0: Option Base 1";

            string expectedCode = "Option Explicit: : Option Base 1";

            //Arrange
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new OptionBaseZeroInspection(parser.State);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(parser.State, CancellationToken.None).Result;

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.First().Fix();
            }

            var actual = module.Content();
            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void OptionBaseZeroStatement_QuickFixWorks_RemoveStatement_InstructionSeparatorAndMultipleLines()
        {
            string inputCode =
@"Option Explicit: Option _
Base _
0: Option Base 1";

            string expectedCode = "Option Explicit: : Option Base 1";

            //Arrange
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new OptionBaseZeroInspection(parser.State);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(parser.State, CancellationToken.None).Result;

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.First().Fix();
            }

            var actual = module.Content();
            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new OptionBaseZeroInspection(null);
            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "OptionBaseZeroInspection";
            var inspection = new OptionBaseZeroInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        private Configuration GetTestConfig()
        {
            var settings = new CodeInspectionSettings();
            settings.CodeInspections.Add(new CodeInspectionSetting
            {
                Description = new OptionBaseInspection(null).Description,
                Severity = CodeInspectionSeverity.Hint
            });
            return new Configuration
            {
                UserSettings = new UserSettings(null, null, null, settings, null, null, null)
            };
        }
    }
}
