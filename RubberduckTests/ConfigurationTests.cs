using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Settings;

namespace RubberduckTests
{
    [TestClass]
    public class ConfigurationTests
    {
        [TestMethod]
        public void GetDefaultTodoMarkersTest()
        {
            var configService = new ConfigurationLoader();

            ToDoMarker[] markers = configService.GetDefaultTodoMarkers();
            Assert.AreEqual("NOTE:", markers[0].Text,"Note failed to load.");
            Assert.AreEqual("TODO:", markers[1].Text,"Todo failed to load.");
            Assert.AreEqual("BUG:", markers[2].Text,"Bug failed to load.");
        }

        [TestMethod]
        public void ToDoMarkersTextIsNotNull()
        {
            var configService = new ConfigurationLoader();
            ToDoMarker[] markers = configService.LoadConfiguration().UserSettings.ToDoListSettings.ToDoMarkers;

            foreach (var marker in markers)
            {
                Assert.IsNotNull(marker.Text);
            }
        }

        [TestMethod]
        public void DefaultCodeInspectionsIsNotNull()
        {
            var configService = new ConfigurationLoader();
            var config = configService.GetDefaultCodeInspections();

            Assert.IsNotNull(config);
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
