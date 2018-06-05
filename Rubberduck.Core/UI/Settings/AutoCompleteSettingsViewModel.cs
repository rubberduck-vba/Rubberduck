using NLog;
using Rubberduck.Resources;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Data;

namespace Rubberduck.UI.Settings
{
    public class AutoCompleteSettingsViewModel : SettingsViewModelBase, ISettingsViewModel
    {
        public AutoCompleteSettingsViewModel(Configuration config)
        {
            Settings = new ListCollectionView(
                config.UserSettings.AutoCompleteSettings.AutoCompletes.ToList());

            ExportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ExportSettings());
            ImportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ImportSettings());
        }

        private ListCollectionView _settings;
        public ListCollectionView Settings
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
            config.UserSettings.AutoCompleteSettings.AutoCompletes = new HashSet<AutoCompleteSetting>(_settings.OfType<AutoCompleteSetting>());
        }

        private void TransferSettingsToView(Rubberduck.Settings.AutoCompleteSettings toLoad)
        {
            Settings = new ListCollectionView(toLoad.AutoCompletes.ToList());
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
                     AutoCompletes = new HashSet<AutoCompleteSetting>(this.Settings.SourceCollection.OfType<AutoCompleteSetting>()),
                });
            }
        }
    }
}
