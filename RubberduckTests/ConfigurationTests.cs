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
            Assert.AreEqual("'NOTE:", markers[0].text,"Note failed to load.");
            Assert.AreEqual("'TODO:", markers[1].text,"Todo failed to load.");
            Assert.AreEqual("'BUG:", markers[2].text,"Bug failed to load.");
        }
    }
}
