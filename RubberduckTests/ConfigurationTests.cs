using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Settings;

namespace RubberduckTests
{
    [TestClass]
    public class ConfigurationTests
    {
        [TestMethod]
        public void GetDefaultTodoMarkersTest()
        {
            var settings = new ToDoListSettings();

            ToDoMarker[] markers = settings.ToDoMarkers;
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
            var actual = new CodeInspectionSettings(new HashSet<CodeInspectionSetting> {new CodeInspectionSetting(inspection.Object)}, new WhitelistedIdentifierSetting[] {}, true).CodeInspections;

            Assert.AreEqual(expected.Length, actual.Count);
            Assert.AreEqual(inspection.Object.Name, actual.First().Name);
            Assert.AreEqual(inspection.Object.Severity, actual.First().Severity);
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
