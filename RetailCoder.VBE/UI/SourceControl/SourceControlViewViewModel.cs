using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Ninject;
using Rubberduck.Settings;
using Rubberduck.SourceControl;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.SourceControl
{
    public class SourceControlViewViewModel : ViewModelBase
    {
        private readonly VBE _vbe;
        private readonly ISourceControlProviderFactory _providerFactory;
        private readonly IFolderBrowserFactory _folderBrowserFactory;
        private readonly IConfigurationService<SourceControlConfiguration> _configService;
        private readonly SourceControlConfiguration _config;
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public SourceControlViewViewModel(
            VBE vbe,
            ISourceControlProviderFactory providerFactory,
            IFolderBrowserFactory folderBrowserFactory,
            IConfigurationService<SourceControlConfiguration> configService,
            [Named("changesView")] IControlView changesView,
            [Named("branchesView")] IControlView branchesView,
            [Named("unsyncedCommitsView")] IControlView unsyncedCommitsView,
            [Named("settingsView")] IControlView settingsView,
            ICodePaneWrapperFactory wrapperFactory)
        {
            _vbe = vbe;
            _providerFactory = providerFactory;
            _folderBrowserFactory = folderBrowserFactory;

            _configService = configService;
            _config = _configService.LoadConfiguration();
            _wrapperFactory = wrapperFactory;

            _initRepoCommand = new DelegateCommand(_ => InitRepo());
            _openRepoCommand = new DelegateCommand(_ => OpenRepo());
            _cloneRepoCommand = new DelegateCommand(_ => ShowCloneRepoGrid());
            _refreshCommand = new DelegateCommand(_ => Refresh());
            _dismissErrorMessageCommand = new DelegateCommand(_ => DismissErrorMessage());
            _showFilePickerCommand = new DelegateCommand(_ => ShowFilePicker());
            _closeLoginGridCommand = new DelegateCommand(_ => CloseLoginGrid());

            _cloneRepoOkButtonCommand = new DelegateCommand(_ => CloneRepo(), _ => !IsNotValidRemotePath);
            _cloneRepoCancelButtonCommand = new DelegateCommand(_ => CloseCloneRepoGrid());

            TabItems = new ObservableCollection<IControlView>
            {
                changesView,
                branchesView,
                unsyncedCommitsView,
                settingsView
            };
            Status = RubberduckUI.Offline;

            ListenForErrors();
        }

        private ISourceControlProvider _provider;
        public ISourceControlProvider Provider
        {
            get { return _provider; }
            set
            {
                _provider = value;
                SetChildPresenterSourceControlProviders(_provider);
            }
        }

        private ObservableCollection<IControlView> _tabItems;
        public ObservableCollection<IControlView> TabItems
        {
            get { return _tabItems; }
            set
            {
                if (_tabItems != value)
                {
                    _tabItems = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _status;
        public string Status
        {
            get { return _status; }
            set
            {
                if (_status != value)
                {
                    _status = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _displayCloneRepoGrid;
        public bool DisplayCloneRepoGrid
        {
            get { return _displayCloneRepoGrid; }
            set
            {
                if (_displayCloneRepoGrid != value)
                {
                    _displayCloneRepoGrid = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _remotePath;
        public string RemotePath
        {
            get { return _remotePath; }
            set
            {
                if (_remotePath != value)
                {
                    _remotePath = value;
                    LocalDirectory =
                        _config.DefaultRepositoryLocation +
                        (_config.DefaultRepositoryLocation.EndsWith("\\") ? string.Empty : "\\") +
                        _remotePath.Split('/').Last().Replace(".git", string.Empty);

                    OnPropertyChanged();
                    OnPropertyChanged("IsNotValidRemotePath");
                }
            }
        }

        private string _localDirectory;
        public string LocalDirectory
        {
            get { return _localDirectory; }
            set
            {
                if (_localDirectory != value)
                {
                    _localDirectory = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _displayErrorMessageGrid;
        public bool DisplayErrorMessageGrid
        {
            get { return _displayErrorMessageGrid; }
            set
            {
                if (_displayErrorMessageGrid != value)
                {
                    _displayErrorMessageGrid = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _errorMessage;
        public string ErrorMessage
        {
            get { return _errorMessage; }
            set
            {
                if (_errorMessage != value)
                {
                    _errorMessage = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IsNotValidRemotePath
        {
            get
            {
                Uri uri;
                return !Uri.TryCreate(RemotePath, UriKind.Absolute, out uri);
            }
        }

        private bool _displayLoginGrid;
        public bool DisplayLoginGrid
        {
            get { return _displayLoginGrid; }
            set
            {
                if (_displayLoginGrid != value)
                {
                    _displayLoginGrid = value;
                    OnPropertyChanged();
                }
            }
        }

        private void ListenForErrors()
        {
            foreach (var tab in TabItems)
            {
                tab.ViewModel.ErrorThrown += ViewModel_ErrorThrown;
            }
        }

        private void ViewModel_ErrorThrown(object sender, ErrorEventArgs e)
        {
            const string unauthorizedMessage = "Request failed with status code: 401";

            if (e.InnerMessage == unauthorizedMessage)
            {
                DisplayLoginGrid = true;
            }
            else
            {
                ErrorMessage = e.Message + ":" + Environment.NewLine + e.InnerMessage;
                DisplayErrorMessageGrid = true;
            }
        }

        private void DismissErrorMessage()
        {
            DisplayErrorMessageGrid = false;
        }

        public void CreateProviderWithCredentials(SecureCredentials credentials)
        {
            Provider = _providerFactory.CreateProvider(_vbe.ActiveVBProject, Provider.CurrentRepository, credentials, _wrapperFactory);
        }

        private void InitRepo()
        {
            using (var folderPicker = _folderBrowserFactory.CreateFolderBrowser((RubberduckUI.SourceControl_CreateNewRepo)))
            {
                if (folderPicker.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                _provider = _providerFactory.CreateProvider(_vbe.ActiveVBProject);
                var repo = _provider.InitVBAProject(folderPicker.SelectedPath);
                Provider = _providerFactory.CreateProvider(_vbe.ActiveVBProject, repo, _wrapperFactory);

                AddRepoToConfig((Repository)repo);
                Status = RubberduckUI.Online;
            }
        }

        private void SetChildPresenterSourceControlProviders(ISourceControlProvider provider)
        {
            foreach (var tab in TabItems)
            {
                tab.ViewModel.Provider = provider;
            }
        }

        private void AddRepoToConfig(Repository repo)
        {
            _config.Repositories.Add(repo);
            _configService.SaveConfiguration(_config);
        }

        private void OpenRepo()
        {
            using (var folderPicker = _folderBrowserFactory.CreateFolderBrowser(RubberduckUI.SourceControl_OpenWorkingDirectory, false))
            {
                if (folderPicker.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                var project = _vbe.ActiveVBProject;
                var repo = new Repository(project.Name, folderPicker.SelectedPath, string.Empty);

                try
                {
                    Provider = _providerFactory.CreateProvider(project, repo, _wrapperFactory);
                }
                catch (SourceControlException ex)
                {
                    ViewModel_ErrorThrown(null, new ErrorEventArgs(ex.Message, ex.InnerException.Message));
                    return;
                }

                AddRepoToConfig(repo);

                Status = RubberduckUI.Online;
            }
        }

        private void CloneRepo()
        {
            try
            {
                _provider = _providerFactory.CreateProvider(_vbe.ActiveVBProject);
                var repo = _provider.Clone(RemotePath, LocalDirectory);
                Provider = _providerFactory.CreateProvider(_vbe.ActiveVBProject, repo, _wrapperFactory);
                AddRepoToConfig(new Repository
                {
                    Name = _vbe.ActiveVBProject.Name,
                    LocalLocation = repo.LocalLocation,
                    RemoteLocation = repo.RemoteLocation
                });
            }
            catch (SourceControlException ex)
            {
                ViewModel_ErrorThrown(null, new ErrorEventArgs(ex.Message, ex.InnerException.Message));
                return;
            }

            CloseCloneRepoGrid();

            SetChildPresenterSourceControlProviders(_provider);
            Status = RubberduckUI.Online;
        }

        private void ShowCloneRepoGrid()
        {
            DisplayCloneRepoGrid = true;
        }

        private void CloseCloneRepoGrid()
        {
            RemotePath = string.Empty;

            DisplayCloneRepoGrid = false;
        }

        private void OpenRepoAssignedToProject()
        {
            if (!ValidRepoExists())
            {
                return;
            }

            try
            {
                Provider = _providerFactory.CreateProvider(_vbe.ActiveVBProject,
                    _config.Repositories.First(repo => repo.Name == _vbe.ActiveVBProject.Name), _wrapperFactory);
                Status = RubberduckUI.Online;
            }
            catch (SourceControlException ex)
            {
                ViewModel_ErrorThrown(null, new ErrorEventArgs(ex.Message, ex.InnerException.Message));
                Status = RubberduckUI.Offline;
            }
        }

        private void Refresh()
        {
            if (Provider == null)
            {
                OpenRepoAssignedToProject();
            }
            else
            {
                foreach (var tab in TabItems)
                {
                    tab.ViewModel.RefreshView();
                }
            }
        }

        private bool ValidRepoExists()
        {
            if (_config.Repositories == null)
            {
                return false;
            }

            var possibleRepos = _config.Repositories.Where(repo => repo.Name == _vbe.ActiveVBProject.Name);

            var possibleCount = possibleRepos.Count();

            //todo: if none are found, prompt user to create one
            //todo: more than one are found, prompt for correct one
            return possibleCount != 0 && possibleCount <= 1;
        }

        private void ShowFilePicker()
        {
            using (var folderPicker = _folderBrowserFactory.CreateFolderBrowser("Default Repository Directory"))
            {
                if (folderPicker.ShowDialog() == DialogResult.OK)
                {
                    LocalDirectory = folderPicker.SelectedPath;
                }
            }
        }

        private void CloseLoginGrid()
        {
            DisplayLoginGrid = false;
        }

        private readonly ICommand _refreshCommand;
        public ICommand RefreshCommand
        {
            get { return _refreshCommand; }
        }

        private readonly ICommand _initRepoCommand;
        public ICommand InitRepoCommand
        {
            get { return _initRepoCommand; }
        }

        private readonly ICommand _openRepoCommand;
        public ICommand OpenRepoCommand
        {
            get { return _openRepoCommand; }
        }

        private readonly ICommand _cloneRepoCommand;
        public ICommand CloneRepoCommand
        {
            get { return _cloneRepoCommand; }
        }
        
        private readonly ICommand _showFilePickerCommand;
        public ICommand ShowFilePickerCommand
        {
            get
            {
                return _showFilePickerCommand;
            }
        }

        private readonly ICommand _cloneRepoOkButtonCommand;
        public ICommand CloneRepoOkButtonCommand
        {
            get
            {
                return _cloneRepoOkButtonCommand;
            }
        }

        private readonly ICommand _cloneRepoCancelButtonCommand;
        public ICommand CloneRepoCancelButtonCommand
        {
            get
            {
                return _cloneRepoCancelButtonCommand;
            }
        }

        private readonly ICommand _dismissErrorMessageCommand;
        public ICommand DismissErrorMessageCommand
        {
            get
            {
                return _dismissErrorMessageCommand;
            }
        }

        private readonly ICommand _closeLoginGridCommand;
        public ICommand CloseLoginGridCommand
        {
            get
            {
                return _closeLoginGridCommand;
            }
        }
    }
}
