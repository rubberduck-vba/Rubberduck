using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;
using Rubberduck.Settings;
using Rubberduck.Inspections.Rubberduck.Inspections;
using System.Threading;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ObsoleteLetStatementInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteLetStatement_ReturnsResult()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    Let var2 = var1
End Sub";

            //Arrange
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteLetStatementInspection(parser.State);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(parser.State, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteLetStatement_ReturnsResult_MultipleLets()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    Let var2 = var1
    Let var1 = var2
End Sub";

            //Arrange
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteLetStatementInspection(parser.State);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(parser.State, CancellationToken.None).Result;

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteLetStatement_DoesNotReturnResult()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    var2 = var1
End Sub";

            //Arrange
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteLetStatementInspection(parser.State);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(parser.State, CancellationToken.None).Result;

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteLetStatement_ReturnsResult_SomeConstantsUsed()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    Let var2 = var1
    var1 = var2
End Sub";

            //Arrange
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteLetStatementInspection(parser.State);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(parser.State, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteLetStatement_QuickFixWorks()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    Let var2 = var1
End Sub";

            const string expectedCode =
@"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    var2 = var1
End Sub";

            //Arrange
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteLetStatementInspection(parser.State);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(parser.State, CancellationToken.None).Result;

            inspectionResults.First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new ObsoleteLetStatementInspection(null);
            Assert.AreEqual(CodeInspectionType.LanguageOpportunities, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ObsoleteLetStatementInspection";
            var inspection = new ObsoleteLetStatementInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        private Configuration GetTestConfig()
        {
            var settings = new CodeInspectionSettings();
            settings.CodeInspections.Add(new CodeInspectionSetting
            {
                Description = new ObsoleteLetStatementInspection(null).Description,
                Severity = CodeInspectionSeverity.Suggestion
            });
            return new Configuration
            {
                UserSettings = new UserSettings(null, null, null, settings, null, null)
            };
        }
    }
}
