using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Settings;
using Rubberduck.ToDoItems;

namespace RubberduckTests
{
    [TestClass]
    public class ConfigurationTests
    {
        [TestMethod]
        public void GetDefaultTodoMarkersTest()
        {
            var configService = new ConfigurationLoader(null);

            ToDoMarker[] markers = configService.GetDefaultTodoMarkers();
            Assert.AreEqual("NOTE:", markers[0].Text,"Note failed to load.");
            Assert.AreEqual("TODO:", markers[1].Text,"Todo failed to load.");
            Assert.AreEqual("BUG:", markers[2].Text,"Bug failed to load.");
        }

        [TestMethod]
        public void DefaultCodeInspectionsIsAsSpecified()
        {
            var inspection = new Mock<IInspection>();
            //inspection.SetupGet(m => m.Description).Returns("TestInspection");
            //inspection.SetupGet(m => m.Name).Returns("TestInspection");
            //inspection.SetupGet(m => m.Severity).Returns(CodeInspectionSeverity.DoNotShow);

            var expected = new[] { inspection.Object };
            var configService = new ConfigurationLoader(expected);

            var actual = configService.GetDefaultCodeInspections();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.AreEqual(inspection.Object.Name, actual[0].Name);
            Assert.AreEqual(inspection.Object.Severity, actual[0].Severity);
        }

        [TestMethod]
        public void ToStringIsAsExpected()
        {
            var expected = "FixMe:";
            var marker = new ToDoMarker(expected, TodoPriority.High);

            Assert.AreEqual(expected, marker.ToString());
        }
    }
}
