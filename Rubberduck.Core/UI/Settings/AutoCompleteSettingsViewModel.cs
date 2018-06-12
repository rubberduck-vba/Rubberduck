using NLog;
using Rubberduck.Resources;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Rubberduck.UI.Settings
{
    public class AutoCompleteSettingsViewModel : SettingsViewModelBase, ISettingsViewModel
    {
        public AutoCompleteSettingsViewModel(Configuration config)
        {
            Settings = new ObservableCollection<AutoCompleteSetting>(config.UserSettings.AutoCompleteSettings.AutoCompletes);

            ExportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ExportSettings());
            ImportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ImportSettings());
        }

        private ObservableCollection<AutoCompleteSetting> _settings;
        public ObservableCollection<AutoCompleteSetting> Settings
        {
            get { return _settings; }
            set
            {
                if (_settings != value)
                {
                    _settings = value;
                    OnPropertyChanged();
                    SelectAll = value.All(e => e.IsEnabled);
                }
            }
        }

        public void SetToDefaults(Configuration config)
        {
            TransferSettingsToView(config.UserSettings.AutoCompleteSettings);
        }

        public void UpdateConfig(Configuration config)
        {
            config.UserSettings.AutoCompleteSettings.CompleteBlockOnTab = CompleteBlockOnTab;
            config.UserSettings.AutoCompleteSettings.AutoCompletes = new HashSet<AutoCompleteSetting>(_settings);
        }

        private void TransferSettingsToView(Rubberduck.Settings.AutoCompleteSettings toLoad)
        {
            CompleteBlockOnTab = toLoad.CompleteBlockOnTab;
            Settings = new ObservableCollection<AutoCompleteSetting>(toLoad.AutoCompletes);
        }

        private bool _completeBlockOnTab;
        public bool CompleteBlockOnTab
        {
            get { return _completeBlockOnTab; }
            set
            {
                if (_completeBlockOnTab != value)
                {
                    _completeBlockOnTab = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _selectAll;
        public bool SelectAll
        {
            get
            {
                return _selectAll;
            }
            set
            {
                if (_selectAll != value)
                {
                    _selectAll = value;
                    foreach (var setting in Settings)
                    {
                        setting.IsEnabled = value;
                    }
                    OnPropertyChanged();
                }
            }
        }

        private void ImportSettings()
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = RubberduckUI.DialogMask_XmlFilesOnly,
                Title = RubberduckUI.DialogCaption_LoadInspectionSettings
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName)) return;
                var service = new XmlPersistanceService<Rubberduck.Settings.AutoCompleteSettings> { FilePath = dialog.FileName };
                var loaded = service.Load(new Rubberduck.Settings.AutoCompleteSettings());
                TransferSettingsToView(loaded);
            }
        }

        private void ExportSettings()
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = RubberduckUI.DialogMask_XmlFilesOnly,
                Title = RubberduckUI.DialogCaption_SaveInspectionSettings
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName)) return;
                var service = new XmlPersistanceService<Rubberduck.Settings.AutoCompleteSettings> { FilePath = dialog.FileName };
                service.Save(new Rubberduck.Settings.AutoCompleteSettings
                {
                    CompleteBlockOnTab = this.CompleteBlockOnTab,
                    AutoCompletes = new HashSet<AutoCompleteSetting>(Settings),
                });
            }
        }
    }
}
