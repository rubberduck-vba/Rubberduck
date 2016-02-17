using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;

namespace RubberduckTests.Settings
{
    [TestClass]
    public class SettingsControlTests
    {
        [TestMethod]
        public void DefaultViewIsGeneralSettings()
        {
            var configLoader = new ConfigurationLoader(new List<IInspection>());
            var viewModel = new SettingsControlViewModel(configLoader);

            Assert.AreEqual(SettingsViews.GeneralSettings, viewModel.SelectedSettingsView.View);
        }

        [TestMethod]
        public void PassedInViewIsSelected()
        {
            var configLoader = new ConfigurationLoader(new List<IInspection>());
            var viewModel = new SettingsControlViewModel(configLoader, SettingsViews.TodoSettings);

            Assert.AreEqual(SettingsViews.TodoSettings, viewModel.SelectedSettingsView.View);
        }
    }
}