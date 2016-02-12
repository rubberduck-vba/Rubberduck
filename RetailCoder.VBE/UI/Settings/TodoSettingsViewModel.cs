using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using Rubberduck.Settings;
using Rubberduck.ToDoItems;
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

        #region Commands

        private ICommand _addTodoCommand;
        public ICommand AddTodoCommand
        {
            get
            {
                if (_addTodoCommand != null)
                {
                    return _addTodoCommand;
                }
                return _addTodoCommand = new DelegateCommand(_ =>
                {
                    TodoSettings.Add(new ToDoMarker("PLACEHOLDER ", TodoPriority.Low));
                });
            }
        }

        private ICommand _deleteTodoCommand;
        public ICommand DeleteTodoCommand
        {
            get
            {
                if (_deleteTodoCommand != null)
                {
                    return _deleteTodoCommand;
                }
                return _deleteTodoCommand = new DelegateCommand(_ =>
                {
                    TodoSettings.Remove(_ as ToDoMarker);

                    // ReSharper disable once ExplicitCallerInfoArgument
                    OnPropertyChanged("TodoSettings");
                });
            }
        }

        #endregion

        public void UpdateConfig(Configuration config)
        {
            config.UserSettings.ToDoListSettings.ToDoMarkers = TodoSettings.ToArray();
        }

        public void SetToDefaults(Configuration config)
        {
            TodoSettings = new ObservableCollection<ToDoMarker>(
                    config.UserSettings.ToDoListSettings.ToDoMarkers);
        }
    }
}