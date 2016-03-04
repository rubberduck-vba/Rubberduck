using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using Rubberduck.SourceControl;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.SourceControl
{
    public class BranchesViewViewModel : ViewModelBase, IControlViewModel
    {
        public BranchesViewViewModel()
        {
            _newBranchCommand = new DelegateCommand(_ => CreateBranch(), _ => Provider != null);
            _mergeBranchCommand = new DelegateCommand(_ => MergeBranch(), _ => Provider != null);

            _createBranchOkButtonCommand = new DelegateCommand(_ => CreateBranchOk(), _ => !IsNotValidBranchName);
            _createBranchCancelButtonCommand = new DelegateCommand(_ => CreateBranchCancel());

            _mergeBranchesOkButtonCommand = new DelegateCommand(_ => MergeBranchOk(), _ => SourceBranch != DestinationBranch);
            _mergeBranchesCancelButtonCommand = new DelegateCommand(_ => MergeBranchCancel());

            _deleteBranchToolbarButtonCommand = new DelegateCommand(branch => DeleteBranch((string)branch));
            _publishBranchToolbarButtonCommand = new DelegateCommand(branch => PublishBranch((string) branch),
                branch =>
                    Provider != null && Provider.Branches.Any(b => b.Name == (string) branch) &&
                    Provider.Branches.First(b => b.Name == (string) branch).IsRemote);
        }

        private ISourceControlProvider _provider;
        public ISourceControlProvider Provider
        {
            get { return _provider; }
            set
            {
                _provider = value;
                LocalBranches = new ObservableCollection<string>(_provider.Branches.Where(b => !b.IsRemote).Select(b => b.Name));
                Branches = new ObservableCollection<string>(_provider.Branches.Select(b => b.Name));

                CurrentBranch = _provider.CurrentBranch.Name;
            }
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

        private ObservableCollection<string> _branches;
        public ObservableCollection<string> Branches
        {
            get { return _branches; }
            set
            {
                if (_branches != value)
                {
                    _branches = value;
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

                    CreateBranchSource = value;

                    try
                    {
                        Provider.Checkout(_currentBranch);
                    }
                    catch (SourceControlException ex)
                    {
                        RaiseErrorEvent(ex.Message, ex.InnerException.Message);
                    }
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

        private string _createBranchSource;
        public string CreateBranchSource
        {
            get { return _createBranchSource; }
            set
            {
                if (_createBranchSource != value)
                {
                    _createBranchSource = value;
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
                    OnPropertyChanged("IsNotValidBranchName");
                }
            }
        }

        public bool IsNotValidBranchName
        {
            get
            {
                // Rules taken from https://www.kernel.org/pub/software/scm/git/docs/git-check-ref-format.html
                var isValidName = !string.IsNullOrEmpty(NewBranchName) &&
                                  !LocalBranches.Contains(NewBranchName) &&
                                  !NewBranchName.Any(char.IsWhiteSpace) &&
                                  !NewBranchName.Contains("..") &&
                                  !NewBranchName.Contains("~") &&
                                  !NewBranchName.Contains("^") &&
                                  !NewBranchName.Contains(":") &&
                                  !NewBranchName.Contains("?") &&
                                  !NewBranchName.Contains("*") &&
                                  !NewBranchName.Contains("[") &&
                                  !NewBranchName.Contains("//") &&
                                  NewBranchName.FirstOrDefault() != '/' &&
                                  NewBranchName.LastOrDefault() != '/' &&
                                  NewBranchName.LastOrDefault() != '.' &&
                                  NewBranchName != "@" &&
                                  !NewBranchName.Contains("@{") &&
                                  !NewBranchName.Contains("\\");

                if (!isValidName)
                {
                    return true;
                }
                foreach (var section in NewBranchName.Split('/'))
                {
                    isValidName = section.FirstOrDefault() != '.' &&
                                  !section.EndsWith(".lock");
                }

                return !isValidName;
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

        private void CreateBranchOk()
        {
            try
            {
                Provider.CreateBranch(CreateBranchSource, NewBranchName);
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException.Message);
            }

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

        private void DeleteBranch(string branch)
        {
            try
            {
                Provider.DeleteBranch(branch);
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException.Message);
            }

            Branches = new ObservableCollection<string>(_provider.Branches.Select(b => b.Name));
        }

        private void PublishBranch(string branch)
        {
            try
            {
                Provider.Publish(branch);
            }
            catch (SourceControlException ex)
            {
                RaiseErrorEvent(ex.Message, ex.InnerException.Message);
            }

            Branches = new ObservableCollection<string>(_provider.Branches.Select(b => b.Name));
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

        private readonly ICommand _deleteBranchToolbarButtonCommand;
        public ICommand DeleteBranchToolbarButtonCommand
        {
            get
            {
                return _deleteBranchToolbarButtonCommand;
            }
        }

        private readonly ICommand _publishBranchToolbarButtonCommand;
        public ICommand PublishBranchToolbarButtonCommand
        {
            get { return _publishBranchToolbarButtonCommand; }
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