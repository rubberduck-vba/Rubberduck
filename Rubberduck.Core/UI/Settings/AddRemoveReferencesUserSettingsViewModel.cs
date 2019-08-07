using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Input;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public class AddRemoveReferencesUserSettingsViewModel : SettingsViewModelBase<ReferenceSettings>, ISettingsViewModel<ReferenceSettings>
    {
        private readonly IConfigurationService<ReferenceSettings> _provider;
        private readonly IFileSystemBrowserFactory _browserFactory;
        private readonly ReferenceSettings _clean;

        public AddRemoveReferencesUserSettingsViewModel(
            IConfigurationService<ReferenceSettings> provider, 
            IFileSystemBrowserFactory browserFactory,
            IConfigurationService<ReferenceSettings> service)
            : base(service)
        {
            _provider = provider;
            _browserFactory = browserFactory;
            _clean = _provider.Read();

            TransferSettingsToView(_clean);

            IncrementRecentReferencesTrackedCommand = new DelegateCommand(null, ExecuteIncrementRecentReferencesTracked, CanExecuteIncrementMaxConcatLines);
            DecrementReferencesTrackedCommand = new DelegateCommand(null, ExecuteDecrementRecentReferencesTracked, CanExecuteDecrementMaxConcatLines);
            BrowseForPathCommand = new DelegateCommand(null, ExecuteBrowseForPath);
            RemoveSelectedPaths = new DelegateCommand(null, ExecuteRemoveSelectedPaths);
        }

        private int _recent;

        public int RecentReferencesTracked
        {
            get => _recent;
            set
            {
                _recent = value;
                OnPropertyChanged();
            }
        }

        public bool FixBrokenReferences { get; set; }
        public bool AddToRecentOnReferenceEvents { get; set; }

        public ObservableCollection<string> ProjectPaths { get; set; }

        public ICommand BrowseForPathCommand { get; }
        public ICommand RemoveSelectedPaths { get; }

        public ICommand IncrementRecentReferencesTrackedCommand { get; }
        private bool CanExecuteIncrementMaxConcatLines(object parameter) => RecentReferencesTracked < ReferenceSettings.RecentTrackingLimit;     
        private void ExecuteIncrementRecentReferencesTracked(object parameter) => RecentReferencesTracked++;

        public ICommand DecrementReferencesTrackedCommand { get; }
        private void ExecuteDecrementRecentReferencesTracked(object parameter) => RecentReferencesTracked--;
        private bool CanExecuteDecrementMaxConcatLines(object parameter) => RecentReferencesTracked > 0;

        private void ExecuteBrowseForPath(object parameter)
        {
            using (var browser = _browserFactory.CreateFolderBrowser(Resources.RubberduckUI.ReferenceSettings_FolderDialogHeader))
            {
                var result = browser.ShowDialog();
                var path = browser.SelectedPath;
                if (result == DialogResult.OK && 
                    !ProjectPaths.Any(existing => existing.Equals(path, StringComparison.OrdinalIgnoreCase)) &&
                    Directory.Exists(path))
                {
                    ProjectPaths.Add(path);
                }
            }          
        }

        private void ExecuteRemoveSelectedPaths(object parameter)
        {
            if (!(parameter is string path))
            {
                return;
            }

            ProjectPaths.Remove(path);
        }

        protected override void TransferSettingsToView(ReferenceSettings loading)
        {
            RecentReferencesTracked = loading.RecentReferencesTracked;
            FixBrokenReferences = loading.FixBrokenReferences;
            AddToRecentOnReferenceEvents = loading.AddToRecentOnReferenceEvents;
            ProjectPaths = new ObservableCollection<string>(loading.ProjectPaths);
        }

        protected override string DialogLoadTitle { get; }

        protected override string DialogSaveTitle { get; }

        private void TransferViewToSettings(ReferenceSettings target)
        {
            target.RecentReferencesTracked = RecentReferencesTracked;
            target.FixBrokenReferences = FixBrokenReferences;
            target.AddToRecentOnReferenceEvents = AddToRecentOnReferenceEvents;
            target.ProjectPaths = new List<string>(ProjectPaths);
            // ReSharper disable once ExplicitCallerInfoArgument
            OnPropertyChanged("ProjectPaths");
        }

        public void UpdateConfig(Configuration config)
        {
            TransferViewToSettings(_clean);
            _provider.Save(_clean);
        }

        public void SetToDefaults(Configuration config)
        {
            var temp = _provider.ReadDefaults();
            var user = new ReferenceSettings(_clean)
            {
                RecentReferencesTracked = temp.RecentReferencesTracked,
                FixBrokenReferences = temp.FixBrokenReferences,
                AddToRecentOnReferenceEvents = temp.AddToRecentOnReferenceEvents,
                ProjectPaths = temp.ProjectPaths
            };
            TransferSettingsToView(user);
        }
    }
}
