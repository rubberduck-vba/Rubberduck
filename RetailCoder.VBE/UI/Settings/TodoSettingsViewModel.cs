using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using NLog;
using Rubberduck.Settings;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public class TodoSettingsViewModel : ViewModelBase, ISettingsViewModel
    {
        public TodoSettingsViewModel(Configuration config)
        {
            TodoSettings = new ObservableCollection<ToDoMarker>(
                    config.UserSettings.ToDoListSettings.ToDoMarkers);
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
                return _deleteTodoCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), value =>
                {
                    TodoSettings.Remove(value as ToDoMarker);
                });
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
    }
}
