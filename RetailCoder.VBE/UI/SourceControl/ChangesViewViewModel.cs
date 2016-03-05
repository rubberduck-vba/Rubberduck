using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using Rubberduck.SourceControl;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.SourceControl
{
    public class ChangesViewViewModel : ViewModelBase, IControlViewModel
    {
        public ChangesViewViewModel()
        {
            _commitCommand = new DelegateCommand(_ => Commit(), _ => !string.IsNullOrEmpty(CommitMessage) && IncludedChanges != null && IncludedChanges.Any());
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
            CurrentBranch = Provider.CurrentBranch.Name;

            var fileStats = Provider.Status().ToList();

            IncludedChanges = new ObservableCollection<IFileStatusEntry>(fileStats.Where(stat => stat.FileStatus.HasFlag(FileStatus.Modified)));
            UntrackedFiles = new ObservableCollection<IFileStatusEntry>(fileStats.Where(stat => stat.FileStatus.HasFlag(FileStatus.Untracked)));
        }

        private void Provider_BranchChanged(object sender, EventArgs e)
        {
            CurrentBranch = Provider.CurrentBranch.Name;
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

        private ObservableCollection<IFileStatusEntry> _excludedChanges;
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
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException.Message);
            }

            CommitMessage = string.Empty;
        }
        
        private readonly ICommand _commitCommand;
        public ICommand CommitCommand
        {
            get
            {
                return _commitCommand;
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