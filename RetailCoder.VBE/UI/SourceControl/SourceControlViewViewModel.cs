﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.SettingsProvider;
using Rubberduck.SourceControl;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using resx = Rubberduck.UI.SourceControl.SourceControl;
// ReSharper disable ExplicitCallerInfoArgument

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
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly ISourceControlProviderFactory _providerFactory;
        private readonly IFolderBrowserFactory _folderBrowserFactory;
        private readonly IConfigProvider<SourceControlSettings> _configService;
        private readonly IMessageBox _messageBox;
        private readonly IEnvironmentProvider _environment;
        private readonly FileSystemWatcher _fileSystemWatcher;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static readonly IEnumerable<string> VbFileExtensions = new[] { "cls", "bas", "frm" };

        private SourceControlSettings _config;

        public SourceControlViewViewModel(
            IVBE vbe,
            RubberduckParserState state,
            ISourceControlProviderFactory providerFactory,
            IFolderBrowserFactory folderBrowserFactory,
            IConfigProvider<SourceControlSettings> configService,
            IEnumerable<IControlView> views,
            IMessageBox messageBox,
            IEnvironmentProvider environment)
        {
            _vbe = vbe;
            _state = state;
            _providerFactory = providerFactory;
            _folderBrowserFactory = folderBrowserFactory;

            _state.StateChanged += _state_StateChanged;

            _configService = configService;
            _config = _configService.Create();
            _messageBox = messageBox;
            _environment = environment;

            _initRepoCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => InitRepo(), _ => _vbe.VBProjects.Count != 0);
            _openRepoCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => OpenRepo(), _ => _vbe.VBProjects.Count != 0);
            _cloneRepoCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ShowCloneRepoGrid(), _ => _vbe.VBProjects.Count != 0);
            _publishRepoCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ShowPublishRepoGrid(), _ => _vbe.VBProjects.Count != 0 && Provider != null);
            _refreshCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => Refresh());
            _dismissErrorMessageCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DismissErrorMessage());
            _showFilePickerCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ShowFilePicker());
            _loginGridOkCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => CloseLoginGrid(), text => !string.IsNullOrEmpty((string)text));
            _loginGridCancelCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => CloseLoginGrid());

            _cloneRepoOkButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => CloneRepo(), _ => !IsNotValidCloneRemotePath);
            _cloneRepoCancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => CloseCloneRepoGrid());

            _publishRepoOkButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => PublishRepo(), _ => !IsNotValidPublishRemotePath);
            _publishRepoCancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ClosePublishRepoGrid());

            _openCommandPromptCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => OpenCommandPrompt());

            
            AddComponentEventHandlers();

            TabItems = new ObservableCollection<IControlView>(views);
            SetTab(SourceControlTab.Changes);

            Status = RubberduckUI.Offline;

            ListenForErrors();

            _fileSystemWatcher = new FileSystemWatcher();
        }

        public void SetTab(SourceControlTab tab)
        {
            Logger.Trace("Setting active tab to {0}", tab);
            SelectedItem = TabItems.First(t => t.ViewModel.Tab == tab);
        }

        #region Event Handling

        private bool _listening = true;

        private void AddComponentEventHandlers()
        {
            VBEditor.SafeComWrappers.VBA.VBProjects.ProjectRemoved += ProjectRemoved;
            VBEditor.SafeComWrappers.VBA.VBComponents.ComponentAdded += ComponentAdded;
            VBEditor.SafeComWrappers.VBA.VBComponents.ComponentRemoved += ComponentRemoved;
            VBEditor.SafeComWrappers.VBA.VBComponents.ComponentRenamed += ComponentRenamed;
        }

        private void RemoveComponentEventHandlers()
        {
            VBEditor.SafeComWrappers.VBA.VBProjects.ProjectRemoved -= ProjectRemoved;
            VBEditor.SafeComWrappers.VBA.VBComponents.ComponentAdded -= ComponentAdded;
            VBEditor.SafeComWrappers.VBA.VBComponents.ComponentRemoved -= ComponentRemoved;
            VBEditor.SafeComWrappers.VBA.VBComponents.ComponentRenamed -= ComponentRenamed;
        }

        private void ComponentAdded(object sender, ComponentEventArgs e)
        {
            if (!_listening || Provider == null || !Provider.HandleVbeSinkEvents) { return; }

            if (e.ProjectId != Provider.CurrentRepository.Id)
            {
                return;
            }

            Logger.Trace("Component {0} added", e.Component.Name);
            var fileStatus = Provider.Status().SingleOrDefault(stat => Path.GetFileNameWithoutExtension(stat.FilePath) == e.Component.Name);
            if (fileStatus != null)
            {
                Provider.AddFile(fileStatus.FilePath);
            }
        }

        private void ComponentRemoved(object sender, ComponentEventArgs e)
        {
            if (!_listening || Provider == null || !Provider.HandleVbeSinkEvents) { return; }

            if (e.ProjectId != Provider.CurrentRepository.Id)
            {
                return;
            }

            Logger.Trace("Component {0] removed", e.Component.Name);
            var fileStatus = Provider.Status().SingleOrDefault(stat => Path.GetFileNameWithoutExtension(stat.FilePath) == e.Component.Name);
            if (fileStatus != null)
            {
                Provider.RemoveFile(fileStatus.FilePath, true);
            }
        }

        private void ComponentRenamed(object sender, ComponentRenamedEventArgs e)
        {
            if (!_listening || Provider == null || !Provider.HandleVbeSinkEvents) { return; }

            if (e.ProjectId != Provider.CurrentRepository.Id)
            {
                return;
            }

            Logger.Trace("Component {0} renamed to {1}", e.OldName, e.Component.Name);
            var fileStatus = Provider.LastKnownStatus().SingleOrDefault(stat => Path.GetFileNameWithoutExtension(stat.FilePath) == e.OldName);
            if (fileStatus != null)
            {
                var directory = Provider.CurrentRepository.LocalLocation;
                var fileExt = "." + Path.GetExtension(fileStatus.FilePath);

                _fileSystemWatcher.EnableRaisingEvents = false;
                File.Move(Path.Combine(directory, fileStatus.FilePath), Path.Combine(directory, e.Component.Name + fileExt));
                _fileSystemWatcher.EnableRaisingEvents = true;

                Provider.RemoveFile(e.OldName + fileExt, false);
                Provider.AddFile(e.Component.Name + fileExt);
            }
        }

        private void ProjectRemoved(object sender, ProjectEventArgs e)
        {
            if (Provider == null || !Provider.HandleVbeSinkEvents)
            {
                return;
            }

            if (e.ProjectId != Provider.CurrentRepository.Id)
            {
                return;
            }

            _fileSystemWatcher.EnableRaisingEvents = false;
            Provider.Status();  // exports files
            ResetView();
        }

        #endregion

        private void ResetView()
        {
            Logger.Trace("Resetting view");

            _provider = null;
            OnPropertyChanged("RepoDoesNotHaveRemoteLocation");
            Status = RubberduckUI.Offline;

            UiDispatcher.InvokeAsync(() =>
            {
                foreach (var tab in _tabItems)
                {
                    tab.ViewModel.ResetView();
                }
            });
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
                Logger.Trace("Provider changed");

                _provider = value;
                OnPropertyChanged("RepoDoesNotHaveRemoteLocation");
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
            if (!VbFileExtensions.Contains(Path.GetExtension(e.Name)))
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
            if (!VbFileExtensions.Contains(Path.GetExtension(e.Name)))
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
            if (!VbFileExtensions.Contains(e.Name.Split('.').Last()))
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
            if (!VbFileExtensions.Contains(Path.GetExtension(e.Name)))
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
                if (DisplayPublishRepoGrid)
                {
                    _displayPublishRepoGrid = false;
                    OnPropertyChanged("DisplayPublishRepoGrid");
                }

                if (_displayCloneRepoGrid != value)
                {
                    _displayCloneRepoGrid = value;

                    OnPropertyChanged();
                }
            }
        }

        private bool _displayPublishRepoGrid;
        public bool DisplayPublishRepoGrid
        {
            get { return _displayPublishRepoGrid; }
            set
            {
                if (DisplayCloneRepoGrid)
                {
                    _displayCloneRepoGrid = false;
                    OnPropertyChanged("DisplayCloneRepoGrid");
                }

                if (_displayPublishRepoGrid != value)
                {
                    _displayPublishRepoGrid = value;
                    OnPropertyChanged();
                }
            }
        }

        private static readonly Regex LocalFileSystemOrNetworkPathRegex = new Regex(@"^([A-Z]:|\\).*");

        private string _cloneRemotePath;
        public string CloneRemotePath
        {
            get { return _cloneRemotePath; }
            set
            {
                if (_cloneRemotePath != value)
                {
                    _cloneRemotePath = value;
                    var delimiter = LocalFileSystemOrNetworkPathRegex.IsMatch(_cloneRemotePath) ? '\\' : '/';
                    LocalDirectory = Path.Combine(_config.DefaultRepositoryLocation, _cloneRemotePath.Split(delimiter).Last().Replace(".git", string.Empty));
                    OnPropertyChanged();
                    OnPropertyChanged("IsNotValidCloneRemotePath");
                }
            }
        }

        private string _publishRemotePath;
        public string PublishRemotePath
        {
            get { return _publishRemotePath; }
            set
            {
                if (_publishRemotePath != value)
                {
                    _publishRemotePath = value;

                    OnPropertyChanged();
                    OnPropertyChanged("IsNotValidPublishRemotePath");
                }
            }
        }

        public bool RepoDoesNotHaveRemoteLocation
        {
            get
            {
                return !(Provider != null && Provider.RepoHasRemoteOrigin());
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
                if (!Equals(_errorIcon, value))
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

        public bool IsNotValidPublishRemotePath
        {
            get { return !IsValidUri(PublishRemotePath); }
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
                Logger.Trace("Requesting login");
                DisplayLoginGrid = true;
            }
            else
            {
                Logger.Trace("Displaying {0} with title '{1}' and message '{2}'", e.NotificationType, e.Title, e.InnerMessage);

                ErrorTitle = e.Title;
                ErrorMessage = e.InnerMessage;

                IconMappings.TryGetValue(e.NotificationType, out _errorIcon);
                OnPropertyChanged("ErrorIcon");

                DisplayErrorMessageGrid = true;
            }

            if (e.InnerMessage == RubberduckUI.SourceControl_UpdateSettingsMessage)
            {
                _config = _configService.Create();
            }
        }

        private void DismissErrorMessage()
        {
            DisplayErrorMessageGrid = false;
        }

        public void CreateProviderWithCredentials(SecureCredentials credentials)
        {
            if (!_isCloning)
            {
                Provider = _providerFactory.CreateProvider(_vbe.ActiveVBProject, Provider.CurrentRepository, credentials);
            }
            else
            {
                CloneRepo(credentials);
            }
        }

        private void InitRepo()
        {
            using (var folderPicker = _folderBrowserFactory.CreateFolderBrowser(RubberduckUI.SourceControl_CreateNewRepo, false, GetDefaultRepoFolderOrDefault()))
            {
                if (folderPicker.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                Logger.Trace("Initializing repo");

                try
                {
                    _provider = _providerFactory.CreateProvider(_vbe.ActiveVBProject);
                    var repo = _provider.InitVBAProject(folderPicker.SelectedPath);
                    Provider = _providerFactory.CreateProvider(_vbe.ActiveVBProject, repo);

                    AddOrUpdateLocalPathConfig((Repository) repo);
                }
                catch (SourceControlException ex)
                {
                    ViewModel_ErrorThrown(this,
                        new ErrorEventArgs(ex.Message, ex.InnerException, NotificationType.Error));
                }
                catch
                {
                    ViewModel_ErrorThrown(this,
                        new ErrorEventArgs(RubberduckUI.SourceControl_UnknownErrorTitle,
                            RubberduckUI.SourceControl_UnknownErrorMessage, NotificationType.Error));
                    throw;
                }

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
            using (var folderPicker = _folderBrowserFactory.CreateFolderBrowser(RubberduckUI.SourceControl_OpenWorkingDirectory, false, GetDefaultRepoFolderOrDefault()))
            {
                if (folderPicker.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                Logger.Trace("Opening existing repo");
                var project = _vbe.ActiveVBProject;
                var repo = new Repository(project.HelpFile, folderPicker.SelectedPath, string.Empty);

                _listening = false;
                try
                {
                    Provider = _providerFactory.CreateProvider(project, repo);
                }
                catch (SourceControlException ex)
                {
                    _listening = true;
                    ViewModel_ErrorThrown(null, new ErrorEventArgs(ex.Message, ex.InnerException, NotificationType.Error));
                    return;
                }
                catch
                {
                    ViewModel_ErrorThrown(this,
                        new ErrorEventArgs(RubberduckUI.SourceControl_UnknownErrorTitle,
                            RubberduckUI.SourceControl_UnknownErrorMessage, NotificationType.Error));
                    throw;
                }

                _listening = true;

                AddOrUpdateLocalPathConfig(repo);

                Status = RubberduckUI.Online;
            }
        }

        private bool _isCloning;
        private void CloneRepo(SecureCredentials credentials = null)
        {
            _isCloning = true;
            _listening = false;

            Logger.Trace("Cloning repo");
            try
            {
                _provider = _providerFactory.CreateProvider(_vbe.ActiveVBProject);
                var repo = _provider.Clone(CloneRemotePath, LocalDirectory, credentials);
                AddOrUpdateLocalPathConfig(new Repository
                {
                    Id = _vbe.ActiveVBProject.HelpFile,
                    LocalLocation = repo.LocalLocation,
                    RemoteLocation = repo.RemoteLocation
                });

                Provider = _providerFactory.CreateProvider(_vbe.ActiveVBProject, repo);
            }
            catch (SourceControlException ex)
            {
                const string unauthorizedMessage = "Request failed with status code: 401";
                if (ex.InnerException != null && ex.InnerException.Message != unauthorizedMessage)
                {
                    _isCloning = false;
                }

                ViewModel_ErrorThrown(this, new ErrorEventArgs(ex.Message, ex.InnerException, NotificationType.Error));
                return;
            }
            catch
            {
                ViewModel_ErrorThrown(this,
                    new ErrorEventArgs(RubberduckUI.SourceControl_UnknownErrorTitle,
                        RubberduckUI.SourceControl_UnknownErrorMessage, NotificationType.Error));
                throw;
            }

            _isCloning = false;
            _listening = true;
            CloseCloneRepoGrid();
            
            Status = RubberduckUI.Online;
        }

        private void PublishRepo()
        {
            if (Provider == null)
            {
                ViewModel_ErrorThrown(null,
                    new ErrorEventArgs(RubberduckUI.SourceControl_PublishRepo_FailureTitle,
                        RubberduckUI.SourceControl_PublishRepo_NoOpenRepo, NotificationType.Error));
                return;
            }

            Logger.Trace("Publishing repo to remote");
            try
            {
                Provider.AddOrigin(PublishRemotePath, Provider.CurrentBranch.Name);
                Provider.Publish(Provider.CurrentBranch.Name);
            }
            catch (SourceControlException ex)
            {
                ViewModel_ErrorThrown(null, new ErrorEventArgs(ex.Message, ex.InnerException, NotificationType.Error));
            }
            catch
            {
                ViewModel_ErrorThrown(this,
                    new ErrorEventArgs(RubberduckUI.SourceControl_UnknownErrorTitle,
                        RubberduckUI.SourceControl_UnknownErrorMessage, NotificationType.Error));
                throw;
            }

            OnPropertyChanged("RepoDoesNotHaveRemoteLocation");
            ClosePublishRepoGrid();
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

        private void ShowPublishRepoGrid()
        {
            DisplayPublishRepoGrid = true;
        }

        private void ClosePublishRepoGrid()
        {
            PublishRemotePath = string.Empty;

            DisplayPublishRepoGrid = false;
        }

        private void OpenCommandPrompt()
        {
            Logger.Trace("Opening command prompt");
            try
            {
                Process.Start(_config.CommandPromptLocation);
            }
            catch
            {
                ViewModel_ErrorThrown(this,
                    new ErrorEventArgs(RubberduckUI.SourceControl_UnknownErrorTitle,
                        RubberduckUI.SourceControl_UnknownErrorMessage, NotificationType.Error));
                throw;
            }
        }

        private void OpenRepoAssignedToProject()
        {
            if (!ValidRepoExists())
            {
                return;
            }

            Logger.Trace("Opening repo assigned to project");
            try
            {
                _listening = false;
                Provider = _providerFactory.CreateProvider(_vbe.ActiveVBProject,
                    _config.Repositories.First(repo => repo.Id == _vbe.ActiveVBProject.HelpFile));
                Status = RubberduckUI.Online;
            }
            catch (SourceControlException ex)
            {
                ViewModel_ErrorThrown(null, new ErrorEventArgs(ex.Message, ex.InnerException, NotificationType.Error));
                Status = RubberduckUI.Offline;

                _config.Repositories.Remove(_config.Repositories.FirstOrDefault(repo => repo.Id == _vbe.ActiveVBProject.HelpFile));
                _configService.Save(_config);
            }
            catch
            {
                ViewModel_ErrorThrown(this,
                    new ErrorEventArgs(RubberduckUI.SourceControl_UnknownErrorTitle,
                        RubberduckUI.SourceControl_UnknownErrorMessage, NotificationType.Error));
                throw;
            }

            _listening = true;
        }

        private void Refresh()
        {
            Logger.Trace("Refreshing display");

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

            var project = _vbe.ActiveVBProject ?? (_vbe.VBProjects.Count == 1 ? _vbe.VBProjects[1] : null);

            if (project != null)
            {
                var possibleRepos = _config.Repositories.Where(repo => repo.Id == _vbe.ActiveVBProject.HelpFile);
                return possibleRepos.Count() == 1;
            }

            ViewModel_ErrorThrown(this,
                new ErrorEventArgs(RubberduckUI.SourceControl_NoActiveProject,
                    RubberduckUI.SourceControl_ActivateProject, NotificationType.Error));
            return false;
        }

        private string GetDefaultRepoFolderOrDefault()
        {
            var settings = _configService.Create();
            var folder = settings.DefaultRepositoryLocation;
            if (string.IsNullOrEmpty(folder))
            {
                try
                {
                    folder = _environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                }
                catch
                {
                    // ignored - empty is fine if the environment call fails.
                }
            }
            return folder;
        }

        private void ShowFilePicker()
        {
            using (var folderPicker = _folderBrowserFactory.CreateFolderBrowser("Default Repository Directory", true, GetDefaultRepoFolderOrDefault()))
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

        private readonly CommandBase _refreshCommand;
        public CommandBase RefreshCommand
        {
            get { return _refreshCommand; }
        }

        private readonly CommandBase _initRepoCommand;
        public CommandBase InitRepoCommand
        {
            get { return _initRepoCommand; }
        }

        private readonly CommandBase _openRepoCommand;
        public CommandBase OpenRepoCommand
        {
            get { return _openRepoCommand; }
        }

        private readonly CommandBase _cloneRepoCommand;
        public CommandBase CloneRepoCommand
        {
            get { return _cloneRepoCommand; }
        }

        private readonly CommandBase _showFilePickerCommand;
        public CommandBase ShowFilePickerCommand
        {
            get
            {
                return _showFilePickerCommand;
            }
        }

        private readonly CommandBase _cloneRepoOkButtonCommand;
        public CommandBase CloneRepoOkButtonCommand
        {
            get
            {
                return _cloneRepoOkButtonCommand;
            }
        }

        private readonly CommandBase _cloneRepoCancelButtonCommand;
        public CommandBase CloneRepoCancelButtonCommand
        {
            get
            {
                return _cloneRepoCancelButtonCommand;
            }
        }

        private readonly CommandBase _publishRepoCommand;
        public CommandBase PublishRepoCommand
        {
            get { return _publishRepoCommand; }
        }

        private readonly CommandBase _publishRepoOkButtonCommand;
        public CommandBase PublishRepoOkButtonCommand
        {
            get
            {
                return _publishRepoOkButtonCommand;
            }
        }

        private readonly CommandBase _publishRepoCancelButtonCommand;
        public CommandBase PublishRepoCancelButtonCommand
        {
            get
            {
                return _publishRepoCancelButtonCommand;
            }
        }

        private readonly CommandBase _openCommandPromptCommand;
        public CommandBase OpenCommandPromptCommand
        {
            get
            {
                return _openCommandPromptCommand;
            }
        }

        private readonly CommandBase _dismissErrorMessageCommand;
        public CommandBase DismissErrorMessageCommand
        {
            get
            {
                return _dismissErrorMessageCommand;
            }
        }

        private readonly CommandBase _loginGridOkCommand;
        public CommandBase LoginGridOkCommand
        {
            get
            {
                return _loginGridOkCommand;
            }
        }

        private readonly CommandBase _loginGridCancelCommand;
        public CommandBase LoginGridCancelCommand
        {
            get
            {
                return _loginGridCancelCommand;
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

            RemoveComponentEventHandlers();
        }
    }
}
