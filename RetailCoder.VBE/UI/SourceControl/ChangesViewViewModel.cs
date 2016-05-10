using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using Rubberduck.SourceControl;
using Rubberduck.UI.Command;
// ReSharper disable ExplicitCallerInfoArgument

namespace Rubberduck.UI.SourceControl
{
    public class ChangesViewViewModel : ViewModelBase, IControlViewModel
    {
        public ChangesViewViewModel()
        {
            _commitCommand = new DelegateCommand(_ => Commit(), _ => !string.IsNullOrEmpty(CommitMessage) && IncludedChanges != null && IncludedChanges.Any());

            _includeChangesToolbarButtonCommand = new DelegateCommand(fileStatusEntry => IncludeChanges((IFileStatusEntry)fileStatusEntry));
            _excludeChangesToolbarButtonCommand = new DelegateCommand(fileStatusEntry => ExcludeChanges((IFileStatusEntry)fileStatusEntry));
            _undoChangesToolbarButtonCommand = new DelegateCommand(fileStatusEntry => UndoChanges((IFileStatusEntry) fileStatusEntry));
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
            try
            {
                var localLocation = Provider.CurrentRepository.LocalLocation.EndsWith("\\")
                    ? Provider.CurrentRepository.LocalLocation
                    : Provider.CurrentRepository.LocalLocation + "\\";

                Provider.Undo(localLocation + fileStatusEntry.FilePath);

                RefreshView();
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException.Message, NotificationType.Error);
            }
        }

        private void Commit()
        {
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
                    Provider.Pull();
                    Provider.Push();
                }

                if (CommitAction == CommitAction.CommitAndPush)
                {
                    Provider.Push();
                }

                RefreshView();
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException.Message, NotificationType.Error);
            }

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

        private void IncludeChanges(IFileStatusEntry fileStatusEntry)
        {
            if (UntrackedFiles.FirstOrDefault(f => f.FilePath == fileStatusEntry.FilePath) != null)
            {
                Provider.AddFile(fileStatusEntry.FilePath);
            }
            else
            {
                ExcludedChanges.Remove(ExcludedChanges.FirstOrDefault(f => f.FilePath == fileStatusEntry.FilePath));
            }

            RefreshView();
        }

        private void ExcludeChanges(IFileStatusEntry fileStatusEntry)
        {
            ExcludedChanges.Add(fileStatusEntry);

            RefreshView();
        }
        
        private readonly ICommand _commitCommand;
        public ICommand CommitCommand
        {
            get { return _commitCommand; }
        }

        private readonly ICommand _undoChangesToolbarButtonCommand;
        public ICommand UndoChangesToolbarButtonCommand
        {
            get { return _undoChangesToolbarButtonCommand; }
        }

        private readonly ICommand _excludeChangesToolbarButtonCommand;
        public ICommand ExcludeChangesToolbarButtonCommand
        {
            get { return _excludeChangesToolbarButtonCommand; }
        }

        private readonly ICommand _includeChangesToolbarButtonCommand;
        public ICommand IncludeChangesToolbarButtonCommand
        {
            get { return _includeChangesToolbarButtonCommand; }
        }

        public event EventHandler<ErrorEventArgs> ErrorThrown;
        private void RaiseErrorEvent(string message, string innerMessage, NotificationType notificationType)
        {
            var handler = ErrorThrown;
            if (handler != null)
            {
                handler(this, new ErrorEventArgs(message, innerMessage, notificationType));
            }
        }
    }
}