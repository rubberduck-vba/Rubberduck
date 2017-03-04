using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using NLog;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public class TodoSettingsViewModel : SettingsViewModelBase, ISettingsViewModel
    {
        public TodoSettingsViewModel(Configuration config)
        {
            TodoSettings = new ObservableCollection<ToDoMarker>(config.UserSettings.ToDoListSettings.ToDoMarkers);
            ExportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ExportSettings());
            ImportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ImportSettings());
        }

        private ObservableCollection<ToDoMarker> _todoSettings;
        public ObservableCollection<ToDoMarker> TodoSettings
        {
            get { return _todoSettings; }
            set
            {
                if (_todoSettings != value)
                {
                    _todoSettings = value;
                    OnPropertyChanged();
                }
            }
        }

        private CommandBase _addTodoCommand;
        public CommandBase AddTodoCommand
        {
            get
            {
                if (_addTodoCommand != null)
                {
                    return _addTodoCommand;
                }
                return _addTodoCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ =>
                {
                    var placeholder = TodoSettings.Count(m => m.Text.StartsWith("PLACEHOLDER")) + 1;
                    TodoSettings.Add(
                        new ToDoMarker(string.Format("PLACEHOLDER{0} ",
                            placeholder == 1 ? string.Empty : placeholder.ToString(CultureInfo.InvariantCulture))));
                });
            }
        }

        private CommandBase _deleteTodoCommand;
        public CommandBase DeleteTodoCommand
        {
            get
            {
                if (_deleteTodoCommand != null)
                {
                    return _deleteTodoCommand;
                }
                return _deleteTodoCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), x => TodoSettings.Remove(x as ToDoMarker));
            }
        }

        public void UpdateConfig(Configuration config)
        {
            config.UserSettings.ToDoListSettings.ToDoMarkers = TodoSettings.Select(m => new ToDoMarker(m.Text.ToUpperInvariant())).Distinct().ToArray();
        }

        public void SetToDefaults(Configuration config)
        {
            TodoSettings = new ObservableCollection<ToDoMarker>(
                    config.UserSettings.ToDoListSettings.ToDoMarkers);
        }

        private void ImportSettings()
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = RubberduckUI.DialogMask_XmlFilesOnly,
                Title = RubberduckUI.DialogCaption_LoadToDoSettings
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName)) return;
                var service = new XmlPersistanceService<ToDoListSettings> { FilePath = dialog.FileName };
                var loaded = service.Load(new ToDoListSettings());
                TodoSettings = new ObservableCollection<ToDoMarker>(loaded.ToDoMarkers);
            }
        }

        private void ExportSettings()
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = RubberduckUI.DialogMask_XmlFilesOnly,
                Title = RubberduckUI.DialogCaption_SaveToDoSettings
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName)) return;
                var service = new XmlPersistanceService<ToDoListSettings> { FilePath = dialog.FileName };
                service.Save(new ToDoListSettings { ToDoMarkers = TodoSettings.Select(m => new ToDoMarker(m.Text.ToUpperInvariant())).Distinct().ToArray() });
            }
        }
    }
}
