using System.Collections.ObjectModel;
using System.Windows.Input;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.SourceControl
{
    public class BranchesViewViewModel : ViewModelBase
    {
        public BranchesViewViewModel()
        {
            _newBranchCommand = new DelegateCommand(_ => CreateBranch());
            _mergeBranchCommand = new DelegateCommand(_ => MergeBranch());
            _deleteBranchCommand = new DelegateCommand(_ => DeleteBranch());
        }

        private ObservableCollection<string> _localBranches;
        public ObservableCollection<string> LocalBranches
        {
            get { return _localBranches; }
            set
            {
                if (_localBranches != value)
                {
                    _localBranches = value;
                    OnPropertyChanged();
                }
            }
        }

        private ObservableCollection<string> _publishedBranches;
        public ObservableCollection<string> PublishedBranches
        {
            get { return _publishedBranches; }
            set
            {
                if (_publishedBranches != value)
                {
                    _publishedBranches = value;
                    OnPropertyChanged();
                }
            }
        }

        private ObservableCollection<string> _unpublishedBranches;
        public ObservableCollection<string> UnpublishedBranches
        {
            get { return _unpublishedBranches; }
            set
            {
                if (_unpublishedBranches != value)
                {
                    _unpublishedBranches = value;
                    OnPropertyChanged();
                }
            }
        }

        private void CreateBranch()
        {
        }

        private void MergeBranch()
        {
        }

        private void DeleteBranch()
        {
        }

        private readonly ICommand _newBranchCommand;
        public ICommand NewBranchCommand
        {
            get
            {
                return _newBranchCommand;
            }
        }

        private readonly ICommand _mergeBranchCommand;
        public ICommand MergeBranchCommand
        {
            get
            {
                return _mergeBranchCommand;
            }
        }

        private readonly ICommand _deleteBranchCommand;
        public ICommand DeleteBranchCommand
        {
            get
            {
                return _deleteBranchCommand;
            }
        }
    }
}