using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using Microsoft.Vbe.Interop;
using Ninject;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.SourceControl;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using resx = Rubberduck.UI.RubberduckUI;

namespace Rubberduck.UI.SourceControl
{
    public enum SourceControlTab
    {
        Changes,
        Branches,
        UnsyncedCommits,
        Settings
    }

    public sealed class SourceControlViewViewModel : ViewModelBase, IDisposable
    {
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly ISourceControlProviderFactory _providerFactory;
        private readonly IFolderBrowserFactory _folderBrowserFactory;
        private readonly ISourceControlConfigProvider _configService;
        private readonly SourceControlSettings _config;
        private readonly ICodePaneWrapperFactory _wrapperFactory;
        private readonly IMessageBox _messageBox;
        private readonly FileSystemWatcher _fileSystemWatcher;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public SourceControlViewViewModel(
            VBE vbe,
            RubberduckParserState state,
            ISourceControlProviderFactory providerFactory,
            IFolderBrowserFactory folderBrowserFactory,
            ISourceControlConfigProvider configService,
            [Named("changesView")] IControlView changesView,
            [Named("branchesView")] IControlView branchesView,
            [Named("unsyncedCommitsView")] IControlView unsyncedCommitsView,
            [Named("settingsView")] IControlView settingsView,
            ICodePaneWrapperFactory wrapperFactory,
            IMessageBox messageBox)
        {
            _vbe = vbe;
            _state = state;
            _providerFactory = providerFactory;
            _folderBrowserFactory = folderBrowserFactory;

            _state.StateChanged += _state_StateChanged;

            _configService = configService;
            _config = _configService.Create();
            _wrapperFactory = wrapperFactory;
            _messageBox = messageBox;

            _initRepoCommand = new DelegateCommand(_ => InitRepo());
            _openRepoCommand = new DelegateCommand(_ => OpenRepo());
            _cloneRepoCommand = new DelegateCommand(_ => ShowCloneRepoGrid());
            _createNewRemoteRepoCommand = new DelegateCommand(_ => ShowCreateNewRemoteRepoGrid());
            _refreshCommand = new DelegateCommand(_ => Refresh());
            _dismissErrorMessageCommand = new DelegateCommand(_ => DismissErrorMessage());
            _showFilePickerCommand = new DelegateCommand(_ => ShowFilePicker());
            _loginGridOkCommand = new DelegateCommand(_ => CloseLoginGrid(), text => !string.IsNullOrEmpty((string)text));
            _loginGridCancelCommand = new DelegateCommand(_ => CloseLoginGrid());

            _cloneRepoOkButtonCommand = new DelegateCommand(_ => CloneRepo(), _ => !IsNotValidCloneRemotePath);
            _cloneRepoCancelButtonCommand = new DelegateCommand(_ => CloseCloneRepoGrid());

            _createNewRemoteRepoOkButtonCommand = new DelegateCommand(_ => CreateNewRemoteRepo(), _ => !IsNotValidCreateNewRemoteRemotePath && IsValidBranchName(RemoteBranchName));
            _createNewRemoteRepoCancelButtonCommand = new DelegateCommand(_ => CloseCreateNewRemoteRepoGrid());

            TabItems = new ObservableCollection<IControlView>
            {
                changesView,
                branchesView,
                unsyncedCommitsView,
                settingsView
            };
            SetTab(SourceControlTab.Changes);

            Status = RubberduckUI.Offline;

            ListenForErrors();

            _fileSystemWatcher = new FileSystemWatcher();
        }

        public void SetTab(SourceControlTab tab)
        {
            SelectedItem = TabItems.First(t => t.ViewModel.Tab == tab);
        }

        public void AddComponent(VBComponent component)
        {
            if (Provider == null || !Provider.NotifyVBAChanges) { return; }

            var fileStatus = Provider.Status().SingleOrDefault(stat => stat.FilePath.Split('.')[0] == component.Name);
            if (fileStatus != null)
            {
                Provider.AddFile(fileStatus.FilePath);
            }
        }

        public void RemoveComponent(VBComponent component)
        {
            if (Provider == null || !Provider.NotifyVBAChanges) { return; }

            var fileStatus = Provider.Status().SingleOrDefault(stat => stat.FilePath.Split('.')[0] == component.Name);
            if (fileStatus != null)
            {
                Provider.RemoveFile(fileStatus.FilePath, true);
            }
        }

        public void RenameComponent(VBComponent component, string oldName)
        {
            if (Provider == null || !Provider.NotifyVBAChanges) { return; }

            var fileStatus = Provider.LastKnownStatus().SingleOrDefault(stat => stat.FilePath.Split('.')[0] == oldName);
            if (fileStatus != null)
            {
                var directory = Provider.CurrentRepository.LocalLocation;
                directory += directory.EndsWith("\\") ? string.Empty : "\\";

                var fileExt = "." + fileStatus.FilePath.Split('.').Last();

                _fileSystemWatcher.EnableRaisingEvents = false;
                File.Move(directory + fileStatus.FilePath, directory + component.Name + fileExt);
                _fileSystemWatcher.EnableRaisingEvents = true;

                Provider.RemoveFile(oldName + fileExt, false);
                Provider.AddFile(component.Name + fileExt);
            }
        }

        private static readonly IDictionary<NotificationType, BitmapImage> IconMappings =
            new Dictionary<NotificationType, BitmapImage>
            {
                { NotificationType.Info, GetImageSource((Bitmap) resx.ResourceManager.GetObject("information", CultureInfo.InvariantCulture))},
                { NotificationType.Error, GetImageSource((Bitmap) resx.ResourceManager.GetObject("cross_circle", CultureInfo.InvariantCulture))}
            };

        private void _state_StateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State == ParserState.Pending)
            {
                UiDispatcher.InvokeAsync(Refresh);
            }
        }

        private ISourceControlProvider _provider;
        public ISourceControlProvider Provider
        {
            get { return _provider; }
            set
            {
                _provider = value;
                SetChildPresenterSourceControlProviders(_provider);

                if (_fileSystemWatcher.Path != LocalDirectory && Directory.Exists(_provider.CurrentRepository.LocalLocation))
                {
                    _fileSystemWatcher.Path = _provider.CurrentRepository.LocalLocation;
                    _fileSystemWatcher.EnableRaisingEvents = true;
                    _fileSystemWatcher.IncludeSubdirectories = true;

                    _fileSystemWatcher.Created += _fileSystemWatcher_Created;
                    _fileSystemWatcher.Deleted += _fileSystemWatcher_Deleted;
                    _fileSystemWatcher.Renamed += _fileSystemWatcher_Renamed;
                    _fileSystemWatcher.Changed += _fileSystemWatcher_Changed;
                }
            }
        }

        private void _fileSystemWatcher_Changed(object sender, FileSystemEventArgs e)
        {
            // the file system filter doesn't support multiple filters
            if (!new[] { "cls", "bas", "frm" }.Contains(e.Name.Split('.').Last()))
            {
                return;
            }

            if (!Provider.NotifyExternalFileChanges)
            {
                return;
            }

            Logger.Trace("File system watcher detected file changed");
            if (_messageBox.Show(RubberduckUI.SourceControl_ExternalModifications, RubberduckUI.SourceControlPanel_Caption,
                MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1) == DialogResult.OK)
            {
                Provider.ReloadComponent(e.Name);
                UiDispatcher.InvokeAsync(Refresh);
            }
        }

        private void _fileSystemWatcher_Renamed(object sender, RenamedEventArgs e)
        {
            // the file system filter doesn't support multiple filters
            if (!new[] { "cls", "bas", "frm" }.Contains(e.Name.Split('.').Last()))
            {
                return;
            }

            if (!Provider.NotifyExternalFileChanges)
            {
                return;
            }

            Logger.Trace("File system watcher detected file renamed");
            if (_messageBox.Show(RubberduckUI.SourceControl_ExternalModifications, RubberduckUI.SourceControlPanel_Caption,
                MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1) == DialogResult.OK)
            {
                Provider.RemoveFile(e.OldFullPath, true);
                Provider.AddFile(e.FullPath);
                UiDispatcher.InvokeAsync(Refresh);
            }
        }

        private void _fileSystemWatcher_Deleted(object sender, FileSystemEventArgs e)
        {
            // the file system filter doesn't support multiple filters
            if (!new[] { "cls", "bas", "frm" }.Contains(e.Name.Split('.').Last()))
            {
                return;
            }

            if (!Provider.NotifyExternalFileChanges)
            {
                return;
            }

            Logger.Trace("File system watcher detected file deleted");
            if (_messageBox.Show(RubberduckUI.SourceControl_ExternalModifications, RubberduckUI.SourceControlPanel_Caption,
                MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1) == DialogResult.OK)
            {
                Provider.RemoveFile(e.FullPath, true);
                UiDispatcher.InvokeAsync(Refresh);
            }
        }

        private void _fileSystemWatcher_Created(object sender, FileSystemEventArgs e)
        {
            // the file system filter doesn't support multiple filters
            if (!new[] { "cls", "bas", "frm" }.Contains(e.Name.Split('.').Last()))
            {
                return;
            }

            if (!Provider.NotifyExternalFileChanges)
            {
                return;
            }

            Logger.Trace("File system watcher detected file created");
            if (_messageBox.Show(RubberduckUI.SourceControl_ExternalModifications, RubberduckUI.SourceControlPanel_Caption,
                MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1) == DialogResult.OK)
            {
                Provider.AddFile(e.FullPath);
                UiDispatcher.InvokeAsync(Refresh);
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

        private IControlView _selectedItem;
        public IControlView SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                if (_selectedItem != value)
                {
                    _selectedItem = value;
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
                if (DisplayCreateNewRemoteRepoGrid)
                {
                    _displayCreateNewRemoteRepoGrid = false;
                    OnPropertyChanged("DisplayCreateNewRemoteRepoGrid");
                }

                if (_displayCloneRepoGrid != value)
                {
                    _displayCloneRepoGrid = value;

                    OnPropertyChanged();
                }
            }
        }

        private bool _displayCreateNewRemoteRepoGrid;
        public bool DisplayCreateNewRemoteRepoGrid
        {
            get { return _displayCreateNewRemoteRepoGrid; }
            set
            {
                if (DisplayCloneRepoGrid)
                {
                    _displayCloneRepoGrid = false;
                    OnPropertyChanged("DisplayCloneRepoGrid");
                }

                if (_displayCreateNewRemoteRepoGrid != value)
                {
                    _displayCreateNewRemoteRepoGrid = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _cloneRemotePath;
        public string CloneRemotePath
        {
            get { return _cloneRemotePath; }
            set
            {
                if (_cloneRemotePath != value)
                {
                    _cloneRemotePath = value;
                    LocalDirectory =
                        _config.DefaultRepositoryLocation +
                        (_config.DefaultRepositoryLocation.EndsWith("\\") ? string.Empty : "\\") +
                        _cloneRemotePath.Split('/').Last().Replace(".git", string.Empty);

                    OnPropertyChanged();
                    OnPropertyChanged("IsNotValidCloneRemotePath");
                }
            }
        }

        private string _createNewRemoteRemotePath;
        public string CreateNewRemoteRemotePath
        {
            get { return _createNewRemoteRemotePath; }
            set
            {
                if (_createNewRemoteRemotePath != value)
                {
                    _createNewRemoteRemotePath = value;

                    OnPropertyChanged();
                    OnPropertyChanged("IsNotValidCreateNewRemoteRemotePath");
                }
            }
        }

        private string _remoteBranchName;
        public string RemoteBranchName
        {
            get { return _remoteBranchName; }
            set
            {
                if (_remoteBranchName != value)
                {
                    _remoteBranchName = value;
                    OnPropertyChanged();
                    OnPropertyChanged("IsNotValidBranchName");
                }
            }
        }

        public bool IsNotValidBranchName
        {
            get
            {
                return !IsValidBranchName(RemoteBranchName);
            }
        }

        public bool IsValidBranchName(string name)
        {
            // Rules taken from https://www.kernel.org/pub/software/scm/git/docs/git-check-ref-format.html
            var isValidName = !string.IsNullOrEmpty(name) &&
                              !name.Any(char.IsWhiteSpace) &&
                              !name.Contains("..") &&
                              !name.Contains("~") &&
                              !name.Contains("^") &&
                              !name.Contains(":") &&
                              !name.Contains("?") &&
                              !name.Contains("*") &&
                              !name.Contains("[") &&
                              !name.Contains("//") &&
                              name.FirstOrDefault() != '/' &&
                              name.LastOrDefault() != '/' &&
                              name.LastOrDefault() != '.' &&
                              name != "@" &&
                              !name.Contains("@{") &&
                              !name.Contains("\\");

            if (!isValidName)
            {
                return false;
            }
            foreach (var section in name.Split('/'))
            {
                isValidName = section.FirstOrDefault() != '.' &&
                              !section.EndsWith(".lock");
            }

            return isValidName;
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

        private string _errorTitle;
        public string ErrorTitle
        {
            get { return _errorTitle; }
            set
            {
                if (_errorTitle != value)
                {
                    _errorTitle = value;
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

        private BitmapImage _errorIcon;
        public BitmapImage ErrorIcon
        {
            get { return _errorIcon; }
            set
            {
                if (_errorIcon != value)
                {
                    _errorIcon = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IsNotValidCloneRemotePath
        {
            get { return !IsValidUri(CloneRemotePath); }
        }

        public bool IsNotValidCreateNewRemoteRemotePath
        {
            get { return !IsValidUri(CreateNewRemoteRemotePath); }
        }

        private bool IsValidUri(string path)
        {
            Uri uri;
            return Uri.TryCreate(path, UriKind.Absolute, out uri);
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
                ErrorTitle = e.Message;
                ErrorMessage = e.InnerMessage;

                IconMappings.TryGetValue(e.NotificationType, out _errorIcon);
                OnPropertyChanged("ErrorIcon");

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

                AddOrUpdateLocalPathConfig((Repository)repo);
                Status = RubberduckUI.Online;
            }
        }

        private void SetChildPresenterSourceControlProviders(ISourceControlProvider provider)
        {
            if (Provider.CurrentBranch == null)
            {
                ViewModel_ErrorThrown(null,
                    new ErrorEventArgs(RubberduckUI.SourceControl_NoBranchesTitle, RubberduckUI.SourceControl_NoBranchesMessage, NotificationType.Error));

                _config.Repositories.Remove(_config.Repositories.FirstOrDefault(repo => repo.Id == _vbe.ActiveVBProject.HelpFile));
                _configService.Save(_config);

                _provider = null;
                Status = RubberduckUI.Offline;
                return;
            }

            foreach (var tab in TabItems)
            {
                tab.ViewModel.Provider = provider;
            }
        }

        private void AddOrUpdateLocalPathConfig(Repository repo)
        {
            if (_config.Repositories.All(repository => repository.LocalLocation != repo.LocalLocation))
            {
                _config.Repositories.Add(repo);
                _configService.Save(_config);
            }
            else
            {
                var existing = _config.Repositories.Single(repository => repository.LocalLocation == repo.LocalLocation);
                if (string.IsNullOrEmpty(repo.RemoteLocation) && !string.IsNullOrEmpty(existing.RemoteLocation))
                {
                    // config already has remote location and correct repository id - nothing to update
                    return;
                }

                existing.Id = repo.Id;
                existing.RemoteLocation = repo.RemoteLocation;

                _configService.Save(_config);
            }
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
                var repo = new Repository(project.HelpFile, folderPicker.SelectedPath, string.Empty);

                OnOpenRepoStarted();
                try
                {
                    Provider = _providerFactory.CreateProvider(project, repo, _wrapperFactory);
                }
                catch (SourceControlException ex)
                {
                    OnOpenRepoCompleted();
                    ViewModel_ErrorThrown(null, new ErrorEventArgs(ex.Message, ex.InnerException.Message, NotificationType.Error));
                    return;
                }
                OnOpenRepoCompleted();

                AddOrUpdateLocalPathConfig(repo);

                Status = RubberduckUI.Online;
            }
        }

        private void CloneRepo()
        {
            OnOpenRepoStarted();

            try
            {
                _provider = _providerFactory.CreateProvider(_vbe.ActiveVBProject);
                var repo = _provider.Clone(CloneRemotePath, LocalDirectory);
                AddOrUpdateLocalPathConfig(new Repository
                {
                    Id = _vbe.ActiveVBProject.HelpFile,
                    LocalLocation = repo.LocalLocation,
                    RemoteLocation = repo.RemoteLocation
                });

                Provider = _providerFactory.CreateProvider(_vbe.ActiveVBProject, repo, _wrapperFactory);
            }
            catch (SourceControlException ex)
            {
                ViewModel_ErrorThrown(null, new ErrorEventArgs(ex.Message, ex.InnerException.Message, NotificationType.Error));
                return;
            }

            OnOpenRepoCompleted();
            CloseCloneRepoGrid();
            
            Status = RubberduckUI.Online;
        }

        private void CreateNewRemoteRepo()
        {
            if (Provider == null)
            {
                ViewModel_ErrorThrown(null,
                    new ErrorEventArgs(RubberduckUI.SourceControl_CreateNewRemoteRepo_FailureTitle,
                        RubberduckUI.SourceControl_CreateNewRemoteRepo_NoOpenRepo, NotificationType.Error));
                return;
            }

            try
            {
                Provider.AddOrigin(CreateNewRemoteRemotePath, RemoteBranchName);
                Provider.Publish(RemoteBranchName);
            }
            catch (SourceControlException ex)
            {
                ViewModel_ErrorThrown(null, new ErrorEventArgs(ex.Message, ex.InnerException.Message, NotificationType.Error));
            }

            CloseCreateNewRemoteRepoGrid();
        }

        private void ShowCloneRepoGrid()
        {
            DisplayCloneRepoGrid = true;
        }

        private void CloseCloneRepoGrid()
        {
            CloneRemotePath = string.Empty;

            DisplayCloneRepoGrid = false;
        }

        private void ShowCreateNewRemoteRepoGrid()
        {
            DisplayCreateNewRemoteRepoGrid = true;
        }

        private void CloseCreateNewRemoteRepoGrid()
        {
            CreateNewRemoteRemotePath = string.Empty;
            RemoteBranchName = string.Empty;

            DisplayCreateNewRemoteRepoGrid = false;
        }

        private void OpenRepoAssignedToProject()
        {
            if (!ValidRepoExists())
            {
                return;
            }

            try
            {
                OnOpenRepoStarted();
                Provider = _providerFactory.CreateProvider(_vbe.ActiveVBProject,
                    _config.Repositories.First(repo => repo.Id == _vbe.ActiveVBProject.HelpFile), _wrapperFactory);
                Status = RubberduckUI.Online;
            }
            catch (SourceControlException ex)
            {
                ViewModel_ErrorThrown(null, new ErrorEventArgs(ex.Message, ex.InnerException.Message, NotificationType.Error));
                Status = RubberduckUI.Offline;

                _config.Repositories.Remove(_config.Repositories.FirstOrDefault(repo => repo.Id == _vbe.ActiveVBProject.HelpFile));
                _configService.Save(_config);
            }

            OnOpenRepoCompleted();
        }

        private void Refresh()
        {
            _fileSystemWatcher.EnableRaisingEvents = false;
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

            if (Provider != null && Directory.Exists(Provider.CurrentRepository.LocalLocation))
            {
                _fileSystemWatcher.EnableRaisingEvents = true;
            }
        }

        private bool ValidRepoExists()
        {
            if (_config.Repositories == null)
            {
                return false;
            }

            var possibleRepos = _config.Repositories.Where(repo => repo.Id == _vbe.ActiveVBProject.HelpFile);
            var possibleCount = possibleRepos.Count();

            //todo: if none are found, prompt user to create one
            return possibleCount == 1;
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

        private readonly ICommand _createNewRemoteRepoCommand;
        public ICommand CreateNewRemoteRepoCommand
        {
            get { return _createNewRemoteRepoCommand; }
        }

        private readonly ICommand _createNewRemoteRepoOkButtonCommand;
        public ICommand CreateNewRemoteRepoOkButtonCommand
        {
            get
            {
                return _createNewRemoteRepoOkButtonCommand;
            }
        }

        private readonly ICommand _createNewRemoteRepoCancelButtonCommand;
        public ICommand CreateNewRemoteRepoCancelButtonCommand
        {
            get
            {
                return _createNewRemoteRepoCancelButtonCommand;
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

        private readonly ICommand _loginGridOkCommand;
        public ICommand LoginGridOkCommand
        {
            get
            {
                return _loginGridOkCommand;
            }
        }

        private readonly ICommand _loginGridCancelCommand;
        public ICommand LoginGridCancelCommand
        {
            get
            {
                return _loginGridCancelCommand;
            }
        }

        public event EventHandler<EventArgs> OpenRepoStarted;
        private void OnOpenRepoStarted()
        {
            var handler = OpenRepoStarted;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        public event EventHandler<EventArgs> OpenRepoCompleted;
        private void OnOpenRepoCompleted()
        {
            var handler = OpenRepoCompleted;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        public void Dispose()
        {
            if (_state != null)
            {
                _state.StateChanged -= _state_StateChanged;
            }

            if (_fileSystemWatcher != null)
            {
                _fileSystemWatcher.Created -= _fileSystemWatcher_Created;
                _fileSystemWatcher.Deleted -= _fileSystemWatcher_Deleted;
                _fileSystemWatcher.Renamed -= _fileSystemWatcher_Renamed;
                _fileSystemWatcher.Changed -= _fileSystemWatcher_Changed;
                _fileSystemWatcher.Dispose();
            }
        }
    }
}
