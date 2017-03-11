using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;

namespace RubberduckTests.Settings
{
    [TestClass]
    public class TodoSettingsTests
    {
        private Configuration GetDefaultConfig()
        {
            var todoSettings = new ToDoListSettings
            {
                ToDoMarkers = new[]
                {
                    new ToDoMarker("NOTE "),
                    new ToDoMarker("TODO "),
                    new ToDoMarker("BUG ")
                }
            };

            var userSettings = new UserSettings(null, null, todoSettings, null, null, null, null);
            return new Configuration(userSettings);
        }

        private Configuration GetNondefaultConfig()
        {
            var todoSettings = new ToDoListSettings
            {
                ToDoMarkers = new[]
                {
                    new ToDoMarker("PLACEHOLDER ")
                }
            };

            var userSettings = new UserSettings(null, null, todoSettings, null, null, null, null);
            return new Configuration(userSettings);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void SaveConfigWorks()
        {
            var customConfig = GetNondefaultConfig();
            var viewModel = new TodoSettingsViewModel(customConfig);

            var config = GetDefaultConfig();
            viewModel.UpdateConfig(config);

            Assert.IsTrue(config.UserSettings.ToDoListSettings.ToDoMarkers.SequenceEqual(viewModel.TodoSettings));
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void SetDefaultsWorks()
        {
            var viewModel = new TodoSettingsViewModel(GetNondefaultConfig());

            var defaultConfig = GetDefaultConfig();
            viewModel.SetToDefaults(defaultConfig);

            Assert.IsTrue(defaultConfig.UserSettings.ToDoListSettings.ToDoMarkers.SequenceEqual(viewModel.TodoSettings));
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void TodoMarkersAreSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new TodoSettingsViewModel(defaultConfig);

            Assert.IsTrue(defaultConfig.UserSettings.ToDoListSettings.ToDoMarkers.SequenceEqual(viewModel.TodoSettings));
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void AddTodoMarker()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new TodoSettingsViewModel(defaultConfig);

            viewModel.AddTodoCommand.Execute(null);
            var todoMarkersList = defaultConfig.UserSettings.ToDoListSettings.ToDoMarkers.ToList();
            todoMarkersList.Add(new ToDoMarker("PLACEHOLDER "));

            Assert.IsTrue(todoMarkersList.SequenceEqual(viewModel.TodoSettings));
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void DeleteTodoMarker()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new TodoSettingsViewModel(defaultConfig);

            viewModel.DeleteTodoCommand.Execute(defaultConfig.UserSettings.ToDoListSettings.ToDoMarkers[0]);
            var todoMarkersList = defaultConfig.UserSettings.ToDoListSettings.ToDoMarkers.ToList();
            todoMarkersList.Remove(defaultConfig.UserSettings.ToDoListSettings.ToDoMarkers[0]);

            Assert.IsTrue(todoMarkersList.SequenceEqual(viewModel.TodoSettings));
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void AddTodoMarker_ReusesAction()
        {
            var viewModel = new TodoSettingsViewModel(GetDefaultConfig());

            var initialAddTodoCommand = viewModel.AddTodoCommand;
            Assert.AreSame(initialAddTodoCommand, viewModel.AddTodoCommand);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void DeleteTodoMarker_ReusesAction()
        {
            var viewModel = new TodoSettingsViewModel(GetDefaultConfig());

            var initialAddTodoCommand = viewModel.DeleteTodoCommand;
            Assert.AreSame(initialAddTodoCommand, viewModel.DeleteTodoCommand);
        }

        //Somewhat related to https://github.com/rubberduck-vba/Rubberduck/issues/1623
        [TestCategory("Settings")]
        [TestMethod]
        public void DuplicateToDoMarkersAreIgnored()
        {
            var actual = new ToDoListSettings
            {
                ToDoMarkers = new[]
                {
                    new ToDoMarker("NOTE "),
                    new ToDoMarker("TODO "),
                    new ToDoMarker("BUG "),
                    new ToDoMarker("PLACEHOLDER "),
                    new ToDoMarker("PLACEHOLDER ")
                }
            };

            var expected = new ToDoListSettings
            {
                ToDoMarkers = new[]
                {
                    new ToDoMarker("NOTE "),
                    new ToDoMarker("TODO "),
                    new ToDoMarker("BUG "),
                    new ToDoMarker("PLACEHOLDER ")
                }
            };

            Assert.IsTrue(actual.ToDoMarkers.SequenceEqual(expected.ToDoMarkers));
        }
    }
}
