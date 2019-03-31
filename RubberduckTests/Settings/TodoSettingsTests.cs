using System.Linq;
using NUnit.Framework;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;

namespace RubberduckTests.Settings
{
    [TestFixture]
    public class TodoSettingsTests
    {
        private Configuration GetDefaultConfig()
        {
            var todoSettings = new ToDoListSettings
            {
                ToDoMarkers = new[]
                {
                    new ToDoMarker("NOTE"),
                    new ToDoMarker("TODO"),
                    new ToDoMarker("BUG")
                }
            };

            var userSettings = new UserSettings(null, null, null, todoSettings, null, null, null, null);
            return new Configuration(userSettings);
        }

        private Configuration GetNondefaultConfig()
        {
            var todoSettings = new ToDoListSettings
            {
                ToDoMarkers = new[]
                {
                    new ToDoMarker("PLACEHOLDER")
                }
            };

            var userSettings = new UserSettings(null, null, null, todoSettings, null, null, null, null);
            return new Configuration(userSettings);
        }

        [Category("Settings")]
        [Test]
        public void SaveConfigWorks()
        {
            var customConfig = GetNondefaultConfig();
            var viewModel = new TodoSettingsViewModel(customConfig, null);

            var config = GetDefaultConfig();
            viewModel.UpdateConfig(config);

            Assert.IsTrue(config.UserSettings.ToDoListSettings.ToDoMarkers.SequenceEqual(viewModel.TodoSettings));
        }

        [Category("Settings")]
        [Test]
        public void SetDefaultsWorks()
        {
            var viewModel = new TodoSettingsViewModel(GetNondefaultConfig(), null);

            var defaultConfig = GetDefaultConfig();
            viewModel.SetToDefaults(defaultConfig);

            Assert.IsTrue(defaultConfig.UserSettings.ToDoListSettings.ToDoMarkers.SequenceEqual(viewModel.TodoSettings));
        }

        [Category("Settings")]
        [Test]
        public void TodoMarkersAreSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new TodoSettingsViewModel(defaultConfig, null);

            Assert.IsTrue(defaultConfig.UserSettings.ToDoListSettings.ToDoMarkers.SequenceEqual(viewModel.TodoSettings));
        }

        [Category("Settings")]
        [Test]
        public void AddTodoMarker()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new TodoSettingsViewModel(defaultConfig, null);

            viewModel.AddTodoCommand.Execute(null);
            var todoMarkersList = defaultConfig.UserSettings.ToDoListSettings.ToDoMarkers.ToList();
            todoMarkersList.Add(new ToDoMarker("PLACEHOLDER"));

            Assert.IsTrue(todoMarkersList.SequenceEqual(viewModel.TodoSettings));
        }

        [Category("Settings")]
        [Test]
        public void DeleteTodoMarker()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new TodoSettingsViewModel(defaultConfig, null);

            viewModel.DeleteTodoCommand.Execute(defaultConfig.UserSettings.ToDoListSettings.ToDoMarkers[0]);
            var todoMarkersList = defaultConfig.UserSettings.ToDoListSettings.ToDoMarkers.ToList();
            todoMarkersList.Remove(defaultConfig.UserSettings.ToDoListSettings.ToDoMarkers[0]);

            Assert.IsTrue(todoMarkersList.SequenceEqual(viewModel.TodoSettings));
        }

        [Category("Settings")]
        [Test]
        public void AddTodoMarker_ReusesAction()
        {
            var viewModel = new TodoSettingsViewModel(GetDefaultConfig(), null);

            var initialAddTodoCommand = viewModel.AddTodoCommand;
            Assert.AreSame(initialAddTodoCommand, viewModel.AddTodoCommand);
        }

        [Category("Settings")]
        [Test]
        public void DeleteTodoMarker_ReusesAction()
        {
            var viewModel = new TodoSettingsViewModel(GetDefaultConfig(), null);

            var initialAddTodoCommand = viewModel.DeleteTodoCommand;
            Assert.AreSame(initialAddTodoCommand, viewModel.DeleteTodoCommand);
        }

        //Somewhat related to https://github.com/rubberduck-vba/Rubberduck/issues/1623
        [Category("Settings")]
        [Test]
        public void DuplicateToDoMarkersAreIgnored()
        {
            var actual = new ToDoListSettings
            {
                ToDoMarkers = new[]
                {
                    new ToDoMarker("NOTE"),
                    new ToDoMarker("TODO"),
                    new ToDoMarker("BUG"),
                    new ToDoMarker("PLACEHOLDER"),
                    new ToDoMarker("PLACEHOLDER")
                }
            };

            var expected = new ToDoListSettings
            {
                ToDoMarkers = new[]
                {
                    new ToDoMarker("NOTE"),
                    new ToDoMarker("TODO"),
                    new ToDoMarker("BUG"),
                    new ToDoMarker("PLACEHOLDER")
                }
            };

            Assert.IsTrue(actual.ToDoMarkers.SequenceEqual(expected.ToDoMarkers));
        }
    }
}
