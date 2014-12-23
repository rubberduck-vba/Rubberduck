using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Config;

namespace RubberduckTests
{
    [TestClass]
    public class ConfigurationTests
    {
        [TestMethod]
        public void GetDefaultTodoMarkersTest()
        {
            ToDoMarker[] markers = ConfigurationLoader.GetDefaultTodoMarkers();
            Assert.AreEqual("NOTE:", markers[0].Text,"Note failed to load.");
            Assert.AreEqual("TODO:", markers[1].Text,"Todo failed to load.");
            Assert.AreEqual("BUG:", markers[2].Text,"Bug failed to load.");
        }

        [TestMethod]
        public void ToDoMarkersTextIsNotNull()
        {
            ToDoMarker[] markers = ConfigurationLoader.LoadConfiguration().UserSettings.ToDoListSettings.ToDoMarkers;

            foreach (var marker in markers)
            {
                Assert.IsNotNull(marker.Text);
            }
        }

        [TestMethod]
        public void DefaultCodeInspectionsIsNotNull()
        {
            var config = ConfigurationLoader.GetDefaultCodeInspections();

            Assert.IsNotNull(config);
        }
    }
}
