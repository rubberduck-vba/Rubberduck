using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Input;
using Rubberduck.Parsing.Settings;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public class IgnoredProjectsSettingsViewModel : SettingsViewModelBase<IgnoredProjectsSettings>, ISettingsViewModel<IgnoredProjectsSettings>
    {
        private readonly IConfigurationService<IgnoredProjectsSettings> _provider;
        private readonly IFileSystemBrowserFactory _browserFactory;
        private readonly IgnoredProjectsSettings _currentSettings;

        public IgnoredProjectsSettingsViewModel(
                    IConfigurationService<IgnoredProjectsSettings> provider,
                    IFileSystemBrowserFactory browserFactory,
                    IConfigurationService<IgnoredProjectsSettings> service)
                    : base(service)
        {
            _provider = provider;
            _browserFactory = browserFactory;
            _currentSettings = _provider.Read();

            TransferSettingsToView(_currentSettings);

            AddIgnoredFileCommand = new DelegateCommand(null, ExecuteBrowseForFile);
            RemoveSelectedProjects = new DelegateCommand(null, ExecuteRemoveSelectedFileNames);
        }

        public ObservableCollection<string> IgnoredProjectPaths { get; set; }

        public ICommand AddIgnoredFileCommand { get; }
        public ICommand RemoveSelectedProjects { get; }

        private void ExecuteBrowseForFile(object parameter)
        {
            using (var browser = _browserFactory.CreateOpenFileDialog())
            {
                var result = browser.ShowDialog();
                var fullFilename = browser.FileName;
                if (result == DialogResult.OK &&
                    !IgnoredProjectPaths.Any(existing => existing.Equals(fullFilename, StringComparison.OrdinalIgnoreCase)))
                {
                    IgnoredProjectPaths.Add(fullFilename);
                }
            }
        }

        private void ExecuteRemoveSelectedFileNames(object parameter)
        {
            switch (parameter)
            {
                case string filename:
                    RemoveFilename(filename);
                    return;
                case string[] filenames:
                    foreach (var name in filenames)
                    {
                        RemoveFilename(name);
                    }
                    return;
                default:
                    return;
            }
        }

        private void RemoveFilename(string filename)
        {
            if (IgnoredProjectPaths.Contains(filename))
            {
                IgnoredProjectPaths.Remove(filename);
            }
        }

        protected override void TransferSettingsToView(IgnoredProjectsSettings loading)
        {
            IgnoredProjectPaths = new ObservableCollection<string>(loading.IgnoredProjectPaths);
        }

        protected override string DialogLoadTitle { get; }

        protected override string DialogSaveTitle { get; }

        private void TransferViewToSettings(IgnoredProjectsSettings target)
        {
            target.IgnoredProjectPaths = new List<string>(IgnoredProjectPaths);
            OnPropertyChanged(nameof(IgnoredProjectPaths));
        }

        public void UpdateConfig(Configuration config)
        {
            TransferViewToSettings(_currentSettings);
            _provider.Save(_currentSettings);
        }

        public void SetToDefaults(Configuration config)
        {
            var temp = _provider.ReadDefaults();
            var user = new IgnoredProjectsSettings
            {
                IgnoredProjectPaths = temp.IgnoredProjectPaths
            };
            TransferSettingsToView(user);
        }
    }
}