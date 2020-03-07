using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.Settings;
using Rubberduck.CodeAnalysis.Settings;

namespace RubberduckTests
{
    [TestFixture]
    public class ConfigurationTests
    {
        [Test]
        [Category("Settings")]
        public void GetDefaultTodoMarkersTest()
        {
            var defaultMarkers = new[] {new ToDoMarker("NOTE"), new ToDoMarker("TODO"), new ToDoMarker("BUG")};
            var settings = new ToDoListSettings(defaultMarkers, null);

            ToDoMarker[] markers = settings.ToDoMarkers;
            Assert.AreEqual("NOTE", markers[0].Text.Trim(),"Note failed to load.");
            Assert.AreEqual("TODO", markers[1].Text.Trim(),"Todo failed to load.");
            Assert.AreEqual("BUG" , markers[2].Text.Trim(),"Bug failed to load.");
        }

        [Test]
        [Category("Settings")]
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

        [Test]
        [Category("Settings")]
        public void ToStringIsAsExpected()
        {
            var expected = "FixMe:";
            var marker = new ToDoMarker(expected);

            Assert.AreEqual(expected, marker.ToString());
        }
    }
}
