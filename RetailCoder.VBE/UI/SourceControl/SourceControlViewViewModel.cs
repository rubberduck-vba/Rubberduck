using System.Collections.ObjectModel;
using System.Windows.Controls;
using System.Windows.Input;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.SourceControl
{
    public class SourceControlViewViewModel : ViewModelBase
    {
        private readonly ChangesView _changesView;
        private readonly BranchesView _branchesView;
        private readonly UnsyncedCommitsView _unsyncedCommitsView;
        private readonly SettingsView _settingsView;

        public SourceControlViewViewModel(ChangesView changesView, BranchesView branchesView, UnsyncedCommitsView unsyncedCommitsView, SettingsView settingsView)
        {
            _changesView = changesView;
            _branchesView = branchesView;
            _unsyncedCommitsView = unsyncedCommitsView;
            _settingsView = settingsView;

            _refreshCommand = new DelegateCommand(_ => Refresh());

            SetDataContexts();

            TabItems = new ObservableCollection<TabItem> {_changesView, _branchesView, _unsyncedCommitsView, _settingsView};
        }

        private void SetDataContexts()
        {
            _changesView.DataContext = new ChangesViewViewModel();
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

        private void Refresh()
        {

        }

        private readonly ICommand _refreshCommand;
        public ICommand RefreshCommand
        {
            get
            {
                return _refreshCommand;
            }
        }

        private readonly ICommand _initRepoCommand;
        public ICommand InitRepoCommand
        {
            get
            {
                return _initRepoCommand;
            }
        }

        private readonly ICommand _openRepoCommand;
        public ICommand OpenRepoCommand
        {
            get
            {
                return _openRepoCommand;
            }
        }

        private readonly ICommand _cloneRepoCommand;
        public ICommand CloneRepoCommand
        {
            get
            {
                return _cloneRepoCommand;
            }
        }
    }
}