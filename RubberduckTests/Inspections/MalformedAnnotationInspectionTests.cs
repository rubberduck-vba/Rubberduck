using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using RubberduckTests.Mocks;
using Rubberduck.Settings;
using System.Threading;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete.Rubberduck.Inspections;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class MalformedAnnotationInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void MalformedAnnotation_ReturnsResult_Folder()
        {
            const string inputCode =
@"'@Folder";

            //Arrange
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new MissingAnnotationArgumentInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MalformedAnnotation_DoesNotReturnResult_Folder()
        {
            const string inputCode =
@"'@Folder ""Foo""";

            //Arrange
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new MissingAnnotationArgumentInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MalformedAnnotation_ReturnsResult_Ignore()
        {
            const string inputCode =
@"'@Ignore";

            //Arrange
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new MissingAnnotationArgumentInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MalformedAnnotation_DoesNotReturnResult_Ignore()
        {
            const string inputCode =
@"'@Ignore ProcedureNotUsedInspection";

            //Arrange
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new MissingAnnotationArgumentInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MalformedAnnotation_ReturnsMultipleResults()
        {
            const string inputCode =
@"'@Folder
'@Ignore";

            //Arrange
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new MissingAnnotationArgumentInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MalformedAnnotation_NoIgnoreQuickFix()
        {
            const string inputCode =
@"'@Folder
'@Ignore";

            //Arrange
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new MissingAnnotationArgumentInspection(state);
            var inspector = new Inspector(settings.Object, new IInspection[] { inspection });

            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            Assert.IsFalse(inspectionResults.ElementAt(0).QuickFixes.Any(q => q is IgnoreOnceQuickFix));
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new MissingAnnotationArgumentInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "MissingAnnotationArgumentInspection";
            var inspection = new MissingAnnotationArgumentInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        private Configuration GetTestConfig()
        {
            var settings = new CodeInspectionSettings();
            settings.CodeInspections.Add(new CodeInspectionSetting
            {
                Description = new MissingAnnotationArgumentInspection(null).Description,
                Severity = CodeInspectionSeverity.Error
            });
            return new Configuration
            {
                UserSettings = new UserSettings(null, null, null, settings, null, null, null)
            };
        }
    }
}
