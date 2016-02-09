using System.Collections.ObjectModel;
using System.Linq;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public class TodoSetting
    {
        public string Priority { get; set; }
        public string Text { get; set; }
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
                        m => new TodoSetting {Text = m.Text, Priority = m.PriorityLabel}));
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
    }
}