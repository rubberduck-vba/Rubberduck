using System;
using System.Collections.ObjectModel;
using NLog;
using Rubberduck.SourceControl;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.SourceControl
{
    public class UnsyncedCommitsViewViewModel : ViewModelBase, IControlViewModel
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public UnsyncedCommitsViewViewModel()
        {
            _fetchCommitsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => FetchCommits(), _ => Provider != null);
            _pullCommitsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => PullCommits(), _ => Provider != null);
            _pushCommitsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => PushCommits(), _ => Provider != null);
            _syncCommitsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => SyncCommits(), _ => Provider != null);
        }

        private ISourceControlProvider _provider;
        public ISourceControlProvider Provider
        {
            get { return _provider; }
            set
            {
                Logger.Trace("Provider changed");

                _provider = value;
                _provider.BranchChanged += Provider_BranchChanged;

                RefreshView();
            }
        }

        public void RefreshView()
        {
            Logger.Trace("Refreshing view");

            CurrentBranch = Provider.CurrentBranch.Name;

            IncomingCommits = new ObservableCollection<ICommit>(Provider.UnsyncedRemoteCommits);
            OutgoingCommits = new ObservableCollection<ICommit>(Provider.UnsyncedLocalCommits);
        }

        public void ResetView()
        {
            Logger.Trace("Resetting view");

            _provider.BranchChanged -= Provider_BranchChanged;
            _provider = null;
            CurrentBranch = string.Empty;

            IncomingCommits = new ObservableCollection<ICommit>();
            OutgoingCommits = new ObservableCollection<ICommit>();
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
                Logger.Trace("Fetching");
                Provider.Fetch();

                RefreshView();
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException, NotificationType.Error);
            }
            catch
            {
                RaiseErrorEvent(RubberduckUI.SourceControl_UnknownErrorTitle,
                    RubberduckUI.SourceControl_UnknownErrorMessage, NotificationType.Error);
                throw;
            }
        }

        private void PullCommits()
        {
            try
            {
                Logger.Trace("Pulling");
                Provider.Pull();

                RefreshView();
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException, NotificationType.Error);
            }
            catch
            {
                RaiseErrorEvent(RubberduckUI.SourceControl_UnknownErrorTitle,
                    RubberduckUI.SourceControl_UnknownErrorMessage, NotificationType.Error);
                throw;
            }
        }

        private void PushCommits()
        {
            try
            {
                Logger.Trace("Pushing");
                Provider.Push();

                RefreshView();
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException, NotificationType.Error);
            }
            catch
            {
                RaiseErrorEvent(RubberduckUI.SourceControl_UnknownErrorTitle,
                    RubberduckUI.SourceControl_UnknownErrorMessage, NotificationType.Error);
                throw;
            }
        }

        private void SyncCommits()
        {
            try
            {
                Logger.Trace("Syncing (pull + push)");
                Provider.Pull();
                Provider.Push();

                RefreshView();
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException, NotificationType.Error);
            }
            catch
            {
                RaiseErrorEvent(RubberduckUI.SourceControl_UnknownErrorTitle,
                    RubberduckUI.SourceControl_UnknownErrorMessage, NotificationType.Error);
                throw;
            }
        }

        private readonly CommandBase _fetchCommitsCommand;
        public CommandBase FetchCommitsCommand
        {
            get
            {
                return _fetchCommitsCommand;
            }
        }

        private readonly CommandBase _pullCommitsCommand;
        public CommandBase PullCommitsCommand
        {
            get
            {
                return _pullCommitsCommand;
            }
        }

        private readonly CommandBase _pushCommitsCommand;
        public CommandBase PushCommitsCommand
        {
            get
            {
                return _pushCommitsCommand;
            }
        }

        private readonly CommandBase _syncCommitsCommand;
        public CommandBase SyncCommitsCommand
        {
            get
            {
                return _syncCommitsCommand;
            }
        }

        public event EventHandler<ErrorEventArgs> ErrorThrown;
        private void RaiseErrorEvent(string message, Exception innerException, NotificationType notificationType)
        {
            var handler = ErrorThrown;
            if (handler != null)
            {
                handler(this, new ErrorEventArgs(message, innerException, notificationType));
            }
        }

        private void RaiseErrorEvent(string title, string message, NotificationType notificationType)
        {
            var handler = ErrorThrown;
            if (handler != null)
            {
                handler(this, new ErrorEventArgs(title, message, notificationType));
            }
        }
    }
}
