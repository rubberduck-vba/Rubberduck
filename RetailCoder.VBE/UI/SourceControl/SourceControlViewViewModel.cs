using System.Collections.ObjectModel;
using System.Windows.Controls;
using System.Windows.Input;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.SourceControl
{
    public class SourceControlViewViewModel : ViewModelBase
    {
        public SourceControlViewViewModel(ChangesView changesView, BranchesView branchesView, UnsyncedCommitsView unsyncedCommitsView, SettingsView settingsView)
        {
            _refreshCommand = new DelegateCommand(_ => Refresh());
            _initRepoCommand = new DelegateCommand(_ => InitRepo());
            _openRepoCommand = new DelegateCommand(_ => OpenRepo());
            _cloneRepoCommand = new DelegateCommand(_ => CloneRepo());

            TabItems = new ObservableCollection<TabItem> {changesView, branchesView, unsyncedCommitsView, settingsView};
        }
        
        private ObservableCollection<TabItem> _tabItems;
        public ObservableCollection<TabItem> TabItems
        {
            get { return _tabItems; }
            set
            {
                if (_tabItems != value)
                {
                    _tabItems = value;
                    OnPropertyChanged();
                }
            }
        }

        private void Refresh() { }

        private void InitRepo() { }

        private void OpenRepo() { }

        private void CloneRepo() { }

        private readonly ICommand _refreshCommand;
        public ICommand RefreshCommand
        {
            get { return _refreshCommand; }
        }

        private readonly ICommand _initRepoCommand;
        public ICommand InitRepoCommand
        {
            get { return _initRepoCommand; }
        }

        private readonly ICommand _openRepoCommand;
        public ICommand OpenRepoCommand
        {
            get { return _openRepoCommand; }
        }

        private readonly ICommand _cloneRepoCommand;
        public ICommand CloneRepoCommand
        {
            get { return _cloneRepoCommand; }
        }
    }
}