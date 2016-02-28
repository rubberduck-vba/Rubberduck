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
            _fetchCommitsCommand = new DelegateCommand(_ => FetchCommits());
            _pullCommitsCommand = new DelegateCommand(_ => PullCommits());
            _pushCommitsCommand = new DelegateCommand(_ => PushCommits());
            _syncCommitsCommand = new DelegateCommand(_ => SyncCommits());
        }

        private ISourceControlProvider _provider;
        public ISourceControlProvider Provider
        {
            get { return _provider; }
            set { _provider = value; }
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

        private void FetchCommits()
        {
        }

        private void PullCommits()
        {
        }

        private void PushCommits()
        {
        }

        private void SyncCommits()
        {
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
    }
}