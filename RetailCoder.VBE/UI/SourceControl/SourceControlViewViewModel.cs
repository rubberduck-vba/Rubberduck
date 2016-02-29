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

        private ISourceControlProvider _provider;

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

            _cloneRepoOkButtonCommand = new DelegateCommand(_ => CloneRepo(), _ => !IsNotValidRemotePath);
            _cloneRepoCancelButtonCommand = new DelegateCommand(_ => CloseCloneRepoGrid());

            TabItems = new ObservableCollection<IControlView> {changesView, branchesView, unsyncedCommitsView, settingsView};
            Status = RubberduckUI.Offline;
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

        public bool IsNotValidRemotePath
        {
            get
            {
                Uri uri;
                return !Uri.TryCreate(RemotePath, UriKind.Absolute, out uri);
            }
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
                _provider = _providerFactory.CreateProvider(_vbe.ActiveVBProject, repo, _wrapperFactory);

                AddRepoToConfig((Repository)repo);

                SetChildPresenterSourceControlProviders(_provider);
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
                    _provider = _providerFactory.CreateProvider(project, repo, _wrapperFactory);
                }
                catch (SourceControlException ex)
                {
                    //ShowSecondaryPanel(ex.Message, ex.InnerException.Message);
                    return;
                }

                AddRepoToConfig(repo);

                SetChildPresenterSourceControlProviders(_provider);
                Status = RubberduckUI.Online;
            }
        }

        private void CloneRepo()
        {
            IRepository repo;
            try
            {
                _provider = _providerFactory.CreateProvider(_vbe.ActiveVBProject);
                repo = _provider.Clone(RemotePath, LocalDirectory);
            }
            catch (SourceControlException ex)
            {
                //ShowSecondaryPanel(ex.Message, ex.InnerException.Message);
                return;
            }

            AddRepoToConfig((Repository)repo);

            CloseCloneRepoGrid();
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

        private void Refresh()
        {
            if (!ValidRepoExists())
            {
                //_view.Status = RubberduckUI.Offline;
                return;
            }

            try
            {
                _provider = _providerFactory.CreateProvider(_vbe.ActiveVBProject,
                    _config.Repositories.First(repo => repo.Name == _vbe.ActiveVBProject.Name), _wrapperFactory);

                SetChildPresenterSourceControlProviders(_provider);
                Status = RubberduckUI.Online;
            }
            catch (SourceControlException ex)
            {
                //todo: report failure to user and prompt to create or browse
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
    }
}