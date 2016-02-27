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
            _createBranchOkButtonCommand = new DelegateCommand(_ => CreateBranchOk());
            _createBranchCancelButtonCommand = new DelegateCommand(_ => CreateBranchCancel());
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

        private bool _displayNewBranchGrid;
        public bool DisplayNewBranchGrid
        {
            get { return _displayNewBranchGrid; }
            set
            {
                if (_displayNewBranchGrid != value)
                {
                    _displayNewBranchGrid = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _newBranchName;
        public string NewBranchName
        {
            get { return _newBranchName; }
            set
            {
                if (_newBranchName != value)
                {
                    _newBranchName = value;
                    OnPropertyChanged();
                }
            }
        }

        private void CreateBranch()
        {
            DisplayNewBranchGrid = true;
            NewBranchName = string.Empty;
        }

        private void MergeBranch()
        {
        }

        private void DeleteBranch()
        {
        }

        private void CreateBranchOk()
        {
            DisplayNewBranchGrid = false;
            NewBranchName = string.Empty;
        }

        private void CreateBranchCancel()
        {
            DisplayNewBranchGrid = false;
            NewBranchName = string.Empty;
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

        private readonly ICommand _createBranchOkButtonCommand;
        public ICommand CreateBranchOkButtonCommand
        {
            get
            {
                return _createBranchOkButtonCommand;
            }
        }

        private readonly ICommand _createBranchCancelButtonCommand;
        public ICommand CreateBranchCancelButtonCommand
        {
            get
            {
                return _createBranchCancelButtonCommand;
            }
        }
    }
}