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
            Assert.AreEqual("'NOTE:", markers[0].Text,"Note failed to load.");
            Assert.AreEqual("'TODO:", markers[1].Text,"Todo failed to load.");
            Assert.AreEqual("'BUG:", markers[2].Text,"Bug failed to load.");
        }
    }
}
