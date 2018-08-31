using NLog;
using Rubberduck.Resources;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System;

namespace Rubberduck.UI.Settings
{
    public class AutoCompleteSettingsViewModel : SettingsViewModelBase, ISettingsViewModel
    {
        public AutoCompleteSettingsViewModel(Configuration config)
        {
            TransferSettingsToView(config.UserSettings.AutoCompleteSettings);
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
                }
            }
        }

        public void SetToDefaults(Configuration config)
        {
            TransferSettingsToView(config.UserSettings.AutoCompleteSettings);
        }

        public void UpdateConfig(Configuration config)
        {
            config.UserSettings.AutoCompleteSettings.IsEnabled = IsEnabled;
            config.UserSettings.AutoCompleteSettings.CompleteBlockOnTab = CompleteBlockOnTab;
            config.UserSettings.AutoCompleteSettings.CompleteBlockOnEnter = CompleteBlockOnEnter;
            config.UserSettings.AutoCompleteSettings.EnableSmartConcat = EnableSmartConcat;
            config.UserSettings.AutoCompleteSettings.AutoCompletes = new HashSet<AutoCompleteSetting>(_settings);
        }

        private void TransferSettingsToView(Rubberduck.Settings.AutoCompleteSettings toLoad)
        {
            IsEnabled = toLoad.IsEnabled;
            CompleteBlockOnTab = toLoad.CompleteBlockOnTab;
            CompleteBlockOnEnter = toLoad.CompleteBlockOnEnter;
            EnableSmartConcat = toLoad.EnableSmartConcat;
            Settings = new ObservableCollection<AutoCompleteSetting>(toLoad.AutoCompletes);
        }

        private bool _isEnabled;

        public bool IsEnabled
        {
            get { return _isEnabled; }
            set
            {
                if (_isEnabled != value)
                {
                    _isEnabled = value;
                    OnPropertyChanged();
                }
            }
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
                    if (!_completeBlockOnTab && !_completeBlockOnEnter)
                    {
                        // one must be enabled...
                        CompleteBlockOnEnter = true;
                    }
                }
            }
        }

        private bool _enableSmartConcat;
        public bool EnableSmartConcat
        {
            get { return _enableSmartConcat; }
            set
            {
                if (_enableSmartConcat != value)
                {
                    _enableSmartConcat = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _completeBlockOnEnter;
        public bool CompleteBlockOnEnter
        {
            get { return _completeBlockOnEnter; }
            set
            {
                if (_completeBlockOnEnter != value)
                {
                    _completeBlockOnEnter = value;
                    OnPropertyChanged();
                    if (!_completeBlockOnTab && !_completeBlockOnEnter)
                    {
                        // one must be enabled...
                        CompleteBlockOnTab = true;
                    }
                }
            }
        }

        private bool? _selectAll;
        public bool? SelectAll
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
                        if (setting.IsEnabled != (value ?? false))
                        {
                            setting.IsEnabled = value ?? false;
                        }
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
