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
            _commitCommand = new DelegateCommand(_ => Commit());
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

                CurrentBranch = Provider.CurrentBranch.Name;

                var fileStats = Provider.Status().ToList();

                IncludedChanges = new ObservableCollection<IFileStatusEntry>(fileStats.Where(stat => stat.FileStatus.HasFlag(FileStatus.Modified)));
                UntrackedFiles = new ObservableCollection<IFileStatusEntry>(fileStats.Where(stat => stat.FileStatus.HasFlag(FileStatus.Untracked)));
            }
        }

        private string _currentBranch;
        public string CurrentBranch
        {
            get { return _currentBranch; }
            set {
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

                CommitMessage = string.Empty;
            }
            catch (SourceControlException ex)
            {
                //RaiseActionFailedEvent(ex);
            }
        }

        private readonly ICommand _commitCommand;
        public ICommand CommitCommand
        {
            get
            {
                return _commitCommand;
            }
        }
    }
}