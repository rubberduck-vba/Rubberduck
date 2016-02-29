using System.Collections.ObjectModel;
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
            _cloneRepoCommand = new DelegateCommand(_ => CloneRepo());

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

        private void CloneRepo() { }

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
    }
}