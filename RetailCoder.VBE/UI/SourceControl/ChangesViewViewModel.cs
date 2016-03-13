using System;
using System.Collections.Generic;
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
            OnPropertyChanged("IncludedChanges");
            OnPropertyChanged("UntrackedFiles");
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

        public IEnumerable<IFileStatusEntry> IncludedChanges
        {
            get
            {
                return Provider == null
                    ? new ObservableCollection<IFileStatusEntry>()
                    : new ObservableCollection<IFileStatusEntry>(
                        Provider.Status()
                            .Where(stat =>
                                    (stat.FileStatus.HasFlag(FileStatus.Modified) ||
                                    stat.FileStatus.HasFlag(FileStatus.Added)) &&
                                    !ExcludedChanges.Select(f => f.FilePath).Contains(stat.FilePath)));
            }
        }

        private readonly ObservableCollection<IFileStatusEntry> _excludedChanges = new ObservableCollection<IFileStatusEntry>();
        public ObservableCollection<IFileStatusEntry> ExcludedChanges
        {
            get { return _excludedChanges; }
        }

        public IEnumerable<IFileStatusEntry> UntrackedFiles
        {
            get
            {
                return Provider == null
                    ? new ObservableCollection<IFileStatusEntry>()
                    : new ObservableCollection<IFileStatusEntry>(
                        Provider.Status().Where(stat => stat.FileStatus.HasFlag(FileStatus.Untracked)));
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
                RaiseErrorEvent(ex.Message, ex.InnerException.Message);
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
                RaiseErrorEvent(ex.Message, ex.InnerException.Message);
            }

            CommitMessage = string.Empty;
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