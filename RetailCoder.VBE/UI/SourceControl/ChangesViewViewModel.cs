using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using NLog;
using Rubberduck.SourceControl;
using Rubberduck.UI.Command;
// ReSharper disable ExplicitCallerInfoArgument

namespace Rubberduck.UI.SourceControl
{
    public class ChangesViewViewModel : ViewModelBase, IControlViewModel
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public ChangesViewViewModel()
        {
            _commitCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => Commit(), _ => !string.IsNullOrEmpty(CommitMessage) && IncludedChanges != null && IncludedChanges.Any());

            _includeChangesToolbarButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), fileStatusEntry => IncludeChanges((IFileStatusEntry)fileStatusEntry));
            _excludeChangesToolbarButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), fileStatusEntry => ExcludeChanges((IFileStatusEntry)fileStatusEntry));
            _undoChangesToolbarButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), fileStatusEntry => UndoChanges((IFileStatusEntry) fileStatusEntry));
        }

        private string _commitMessage;
        public string CommitMessage
        {
            get { return _commitMessage; }
            set
            {
                if (_commitMessage != value)
                {
                    _commitMessage = value;
                    OnPropertyChanged();
                }
            }
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
            Logger.Trace("Refreshing view");

            OnPropertyChanged("CurrentBranch");

            IncludedChanges = Provider == null
                ? new ObservableCollection<IFileStatusEntry>()
                : new ObservableCollection<IFileStatusEntry>(
                    Provider.Status()
                        .Where(
                            stat =>
                                (stat.FileStatus.HasFlag(FileStatus.Modified) ||
                                 stat.FileStatus.HasFlag(FileStatus.Added) ||
                                 stat.FileStatus.HasFlag(FileStatus.Removed) ||
                                 stat.FileStatus.HasFlag(FileStatus.RenamedInIndex) ||
                                 stat.FileStatus.HasFlag(FileStatus.RenamedInWorkDir)) &&
                                !ExcludedChanges.Select(f => f.FilePath).Contains(stat.FilePath)));

            UntrackedFiles = Provider == null
                ? new ObservableCollection<IFileStatusEntry>()
                : new ObservableCollection<IFileStatusEntry>(
                    Provider.Status().Where(stat => stat.FileStatus.HasFlag(FileStatus.Untracked)));
        }

        public void ResetView()
        {
            Logger.Trace("Resetting view");

            _provider.BranchChanged -= Provider_BranchChanged;
            _provider = null;

            OnPropertyChanged("CurrentBranch");

            IncludedChanges = new ObservableCollection<IFileStatusEntry>();
            UntrackedFiles = new ObservableCollection<IFileStatusEntry>();
        }

        public SourceControlTab Tab { get { return SourceControlTab.Changes; } }

        private void Provider_BranchChanged(object sender, EventArgs e)
        {
            OnPropertyChanged("CurrentBranch");
        }

        public string CurrentBranch
        {
            get { return Provider == null ? string.Empty : Provider.CurrentBranch.Name; }
        }

        public CommitAction CommitAction { get; set; }

        private ObservableCollection<IFileStatusEntry> _includedChanges;
        public ObservableCollection<IFileStatusEntry> IncludedChanges
        {
            get { return _includedChanges; }
            set
            {
                if (_includedChanges != value)
                {
                    _includedChanges = value;
                    OnPropertyChanged();
                }
            }
        }

        private ObservableCollection<IFileStatusEntry> _excludedChanges = new ObservableCollection<IFileStatusEntry>();
        public ObservableCollection<IFileStatusEntry> ExcludedChanges
        {
            get { return _excludedChanges; }
            set 
            {
                if (_excludedChanges != value)
                {
                    _excludedChanges = value;
                    OnPropertyChanged();
                } 
            }
        }

        private ObservableCollection<IFileStatusEntry> _untrackedFiles;
        public ObservableCollection<IFileStatusEntry> UntrackedFiles
        {
            get { return _untrackedFiles; }
            set 
            {
                if (_untrackedFiles != value)
                {
                    _untrackedFiles = value;
                    OnPropertyChanged();
                } 
            }
        }

        private void UndoChanges(IFileStatusEntry fileStatusEntry)
        {
            Logger.Trace("Undoing changes to file {0}", fileStatusEntry.FilePath);

            try
            {
                var file = Path.GetFileName(fileStatusEntry.FilePath);
                Debug.Assert(!string.IsNullOrEmpty(file));
                Provider.Undo(Path.Combine(Provider.CurrentRepository.LocalLocation, file));
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

        private void Commit()
        {
            Logger.Trace("Committing");

            var changes = IncludedChanges.Select(c => c.FilePath).ToList();
            if (!changes.Any())
            {
                return;
            }

            try
            {
                Provider.Stage(changes);
                Provider.Commit(CommitMessage);

                if (CommitAction == CommitAction.CommitAndSync)
                {
                    Logger.Trace("Commit and sync (pull + push)");
                    Provider.Pull();
                    Provider.Push();
                }

                if (CommitAction == CommitAction.CommitAndPush)
                {
                    Logger.Trace("Commit and push");
                    Provider.Push();
                }

                RefreshView();

                switch (CommitAction)
                {
                    case CommitAction.Commit:
                        RaiseErrorEvent(RubberduckUI.SourceControl_CommitStatus, RubberduckUI.SourceControl_CommitStatus_CommitSuccess, NotificationType.Info);
                        return;
                    case CommitAction.CommitAndPush:
                        RaiseErrorEvent(RubberduckUI.SourceControl_CommitStatus, RubberduckUI.SourceControl_CommitStatus_CommitAndPushSuccess, NotificationType.Info);
                        return;
                    case CommitAction.CommitAndSync:
                        RaiseErrorEvent(RubberduckUI.SourceControl_CommitStatus, RubberduckUI.SourceControl_CommitStatus_CommitAndSyncSuccess, NotificationType.Info);
                        return;
                }
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

            CommitMessage = string.Empty;
        }

        private void IncludeChanges(IFileStatusEntry fileStatusEntry)
        {
            if (UntrackedFiles.FirstOrDefault(f => f.FilePath == fileStatusEntry.FilePath) != null)
            {
                Logger.Trace("Tracking file {0}", fileStatusEntry.FilePath);
                Provider.AddFile(fileStatusEntry.FilePath);
            }
            else
            {
                Logger.Trace("Removing file {0} from excluded changes", fileStatusEntry.FilePath);
                ExcludedChanges.Remove(ExcludedChanges.FirstOrDefault(f => f.FilePath == fileStatusEntry.FilePath));
            }

            RefreshView();
        }

        private void ExcludeChanges(IFileStatusEntry fileStatusEntry)
        {
            Logger.Trace("Adding file {0} to excluded changes", fileStatusEntry.FilePath);
            ExcludedChanges.Add(fileStatusEntry);

            RefreshView();
        }
        
        private readonly CommandBase _commitCommand;
        public CommandBase CommitCommand
        {
            get { return _commitCommand; }
        }

        private readonly CommandBase _undoChangesToolbarButtonCommand;
        public CommandBase UndoChangesToolbarButtonCommand
        {
            get { return _undoChangesToolbarButtonCommand; }
        }

        private readonly CommandBase _excludeChangesToolbarButtonCommand;
        public CommandBase ExcludeChangesToolbarButtonCommand
        {
            get { return _excludeChangesToolbarButtonCommand; }
        }

        private readonly CommandBase _includeChangesToolbarButtonCommand;
        public CommandBase IncludeChangesToolbarButtonCommand
        {
            get { return _includeChangesToolbarButtonCommand; }
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
