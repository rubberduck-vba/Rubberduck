using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Settings;

namespace RubberduckTests
{
    [TestClass]
    public class ConfigurationTests
    {
        [TestMethod]
        public void GetDefaultTodoMarkersTest()
        {
            var configService = new ConfigurationLoader(null, null);

            ToDoMarker[] markers = configService.GetDefaultTodoMarkers();
            Assert.AreEqual("NOTE", markers[0].Text.Trim(),"Note failed to load.");
            Assert.AreEqual("TODO", markers[1].Text.Trim(),"Todo failed to load.");
            Assert.AreEqual("BUG" , markers[2].Text.Trim(),"Bug failed to load.");
        }

        [TestMethod]
        public void DefaultCodeInspectionsIsAsSpecified()
        {
            var inspection = new Mock<IInspection>();
            inspection.SetupGet(m => m.Description).Returns("TestInspection");
            inspection.SetupGet(m => m.Name).Returns("TestInspection");
            inspection.SetupGet(m => m.Severity).Returns(CodeInspectionSeverity.DoNotShow);

            var expected = new[] { inspection.Object };
            var configService = new ConfigurationLoader(expected, null);

            var actual = configService.GetDefaultCodeInspections();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.AreEqual(inspection.Object.Name, actual[0].Name);
            Assert.AreEqual(inspection.Object.Severity, actual[0].Severity);
        }

        [TestMethod]
        public void ToStringIsAsExpected()
        {
            var expected = "FixMe:";
            var marker = new ToDoMarker(expected);

            Assert.AreEqual(expected, marker.ToString());
        }
    }
}
