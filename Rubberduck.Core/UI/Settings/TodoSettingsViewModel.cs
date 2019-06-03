using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using NLog;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;
using Rubberduck.Resources.Settings;

namespace Rubberduck.UI.Settings
{
    public sealed class TodoSettingsViewModel : SettingsViewModelBase<ToDoListSettings>, ISettingsViewModel<ToDoListSettings>
    {
        public TodoSettingsViewModel(Configuration config, IConfigurationService<ToDoListSettings> service) 
            : base(service)
        {
            TodoSettings = new ObservableCollection<ToDoMarker>(config.UserSettings.ToDoListSettings.ToDoMarkers);
            ExportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(),
                _ => ExportSettings(new ToDoListSettings
                {
                    ToDoMarkers = TodoSettings.Select(m => new ToDoMarker(m.Text.ToUpperInvariant())).Distinct()
                        .ToArray()
                }));
            ImportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ImportSettings());
        }

        private ObservableCollection<ToDoMarker> _todoSettings;
        public ObservableCollection<ToDoMarker> TodoSettings
        {
            get => _todoSettings;
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
                        new ToDoMarker(
                            $"PLACEHOLDER{(placeholder == 1 ? string.Empty : placeholder.ToString(CultureInfo.InvariantCulture))}"));
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

        protected override string DialogLoadTitle => SettingsUI.DialogCaption_LoadToDoSettings;
        protected override string DialogSaveTitle => SettingsUI.DialogCaption_SaveToDoSettings;
        protected override void TransferSettingsToView(ToDoListSettings toLoad)
        {
            TodoSettings = new ObservableCollection<ToDoMarker>(toLoad.ToDoMarkers);
        }
    }
}
