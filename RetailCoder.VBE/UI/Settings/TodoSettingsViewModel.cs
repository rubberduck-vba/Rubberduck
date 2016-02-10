using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using Rubberduck.Settings;
using Rubberduck.ToDoItems;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public class TodoSetting
    {
        public TodoPriority Priority { get; set; }
        public string Text { get; set; }

        public TodoSetting(ToDoMarker marker)
        {
            Priority = marker.Priority;
            Text = marker.Text;
        }
    }

    public class TodoSettingsViewModel : ViewModelBase
    {
        private readonly IGeneralConfigService _configService;
        private readonly Configuration _config;

        public TodoSettingsViewModel(IGeneralConfigService configService)
        {
            _configService = configService;
            _config = configService.LoadConfiguration();

            TodoSettings = new ObservableCollection<TodoSetting>(
                    _config.UserSettings.ToDoListSettings.ToDoMarkers.Select(
                        m => new TodoSetting(m)));
        }

        private ObservableCollection<TodoSetting> _todoSettings;
        public ObservableCollection<TodoSetting> TodoSettings
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
                    TodoSettings.Add(new TodoSetting(new ToDoMarker("PLACEHOLDER ", TodoPriority.Low)));
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
                    TodoSettings.Remove(_ as TodoSetting);

                    // ReSharper disable once ExplicitCallerInfoArgument
                    OnPropertyChanged("TodoSettings");
                });
            }
        }

        #endregion
    }
}