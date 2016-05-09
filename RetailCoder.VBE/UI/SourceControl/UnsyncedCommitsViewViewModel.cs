using System;
using System.Collections.ObjectModel;
using System.Windows.Input;
using Rubberduck.SourceControl;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.SourceControl
{
    public class UnsyncedCommitsViewViewModel : ViewModelBase, IControlViewModel
    {
        public UnsyncedCommitsViewViewModel()
        {
            _fetchCommitsCommand = new DelegateCommand(_ => FetchCommits(), _ => Provider != null);
            _pullCommitsCommand = new DelegateCommand(_ => PullCommits(), _ => Provider != null);
            _pushCommitsCommand = new DelegateCommand(_ => PushCommits(), _ => Provider != null);
            _syncCommitsCommand = new DelegateCommand(_ => SyncCommits(), _ => Provider != null);
        }

        private ISourceControlProvider _provider;
        public ISourceControlProvider Provider
        {
            get { return _provider; }
            set
            {
                _provider = value;
                _provider.BranchChanged += Provider_BranchChanged;

                RefreshView();
            }
        }

        public void RefreshView()
        {
            CurrentBranch = Provider.CurrentBranch.Name;

            IncomingCommits = new ObservableCollection<ICommit>(Provider.UnsyncedRemoteCommits);
            OutgoingCommits = new ObservableCollection<ICommit>(Provider.UnsyncedLocalCommits);
        }

        public SourceControlTab Tab { get { return SourceControlTab.UnsyncedCommits; } }

        private void Provider_BranchChanged(object sender, EventArgs e)
        {
            CurrentBranch = Provider.CurrentBranch.Name;
        }

        private ObservableCollection<ICommit> _incomingCommits;
        public ObservableCollection<ICommit> IncomingCommits
        {
            get { return _incomingCommits; }
            set
            {
                if (_incomingCommits != value)
                {
                    _incomingCommits = value;
                    OnPropertyChanged();
                }
            }
        }

        private ObservableCollection<ICommit> _outgoingCommits;
        public ObservableCollection<ICommit> OutgoingCommits
        {
            get { return _outgoingCommits; }
            set
            {
                if (_outgoingCommits != value)
                {
                    _outgoingCommits = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _currentBranch;
        public string CurrentBranch
        {
            get { return _currentBranch; }
            set
            {
                if (_currentBranch != value)
                {
                    _currentBranch = value;
                    OnPropertyChanged();
                }
            }
        }

        private void FetchCommits()
        {
            try
            {
                Provider.Fetch();

                RefreshView();
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException.Message);
            }
        }

        private void PullCommits()
        {
            try
            {
                Provider.Pull();

                RefreshView();
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException.Message);
            }
        }

        private void PushCommits()
        {
            try
            {
                Provider.Push();

                RefreshView();
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException.Message);
            }
        }

        private void SyncCommits()
        {
            try
            {
                Provider.Pull();
                Provider.Push();

                RefreshView();
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException.Message);
            }
        }

        private readonly ICommand _fetchCommitsCommand;
        public ICommand FetchCommitsCommand
        {
            get
            {
                return _fetchCommitsCommand;
            }
        }

        private readonly ICommand _pullCommitsCommand;
        public ICommand PullCommitsCommand
        {
            get
            {
                return _pullCommitsCommand;
            }
        }

        private readonly ICommand _pushCommitsCommand;
        public ICommand PushCommitsCommand
        {
            get
            {
                return _pushCommitsCommand;
            }
        }

        private readonly ICommand _syncCommitsCommand;
        public ICommand SyncCommitsCommand
        {
            get
            {
                return _syncCommitsCommand;
            }
        }

        public event EventHandler<ErrorEventArgs> ErrorThrown;
        private void RaiseErrorEvent(string message, string innerMessage)
        {
            var handler = ErrorThrown;
            if (handler != null)
            {
                handler(this, new ErrorEventArgs(message, innerMessage));
            }
        }
    }
}