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

            _mergeBranchesOkButtonCommand = new DelegateCommand(_ => MergeBranchOk());
            _mergeBranchesCancelButtonCommand = new DelegateCommand(_ => MergeBranchCancel());
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

        private bool _displayCreateBranchGrid;
        public bool DisplayCreateBranchGrid
        {
            get { return _displayCreateBranchGrid; }
            set
            {
                if (_displayCreateBranchGrid != value)
                {
                    _displayCreateBranchGrid = value;
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

        private bool _displayMergeBranchesGrid;
        public bool DisplayMergeBranchesGrid
        {
            get { return _displayMergeBranchesGrid; }
            set
            {
                if (_displayMergeBranchesGrid != value)
                {
                    _displayMergeBranchesGrid = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _sourceBranch;
        public string SourceBranch
        {
            get { return _sourceBranch; }
            set
            {
                if (_sourceBranch != value)
                {
                    _sourceBranch = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _destinationBranch;
        public string DestinationBranch
        {
            get { return _destinationBranch; }
            set
            {
                if (_destinationBranch != value)
                {
                    _destinationBranch = value;
                    OnPropertyChanged();
                }
            }
        }

        private void CreateBranch()
        {
            DisplayMergeBranchesGrid = false;

            DisplayCreateBranchGrid = true;
            NewBranchName = string.Empty;
        }

        private void MergeBranch()
        {
            DisplayCreateBranchGrid = false;

            DisplayMergeBranchesGrid = true;
        }

        private void DeleteBranch()
        {
        }

        private void CreateBranchOk()
        {
            DisplayCreateBranchGrid = false;
            NewBranchName = string.Empty;
        }

        private void CreateBranchCancel()
        {
            DisplayCreateBranchGrid = false;
            NewBranchName = string.Empty;
        }

        private void MergeBranchOk()
        {
            DisplayMergeBranchesGrid = false;
        }

        private void MergeBranchCancel()
        {
            DisplayMergeBranchesGrid = false;
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

        private readonly ICommand _mergeBranchesOkButtonCommand;
        public ICommand MergeBranchesOkButtonCommand
        {
            get
            {
                return _mergeBranchesOkButtonCommand;
            }
        }

        private readonly ICommand _mergeBranchesCancelButtonCommand;
        public ICommand MergeBranchesCancelButtonCommand
        {
            get
            {
                return _mergeBranchesCancelButtonCommand;
            }
        }
    }
}