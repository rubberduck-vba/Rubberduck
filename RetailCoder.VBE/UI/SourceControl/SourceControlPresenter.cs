using System;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Settings;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public class SourceControlPresenter : DockablePresenterBase
    {
        private readonly IChangesPresenter _changesPresenter;
        private readonly IBranchesPresenter _branchesPresenter;
        private readonly ISettingsPresenter _settingsPresenter;
        private readonly IFolderBrowserFactory _folderBrowserFactory;
        private readonly ISourceControlProviderFactory _providerFactory;
        private readonly ISourceControlView _view;
        private readonly IConfigurationService<SourceControlConfiguration> _configService;
        private SourceControlConfiguration _config;

        private ISourceControlProvider _provider;

        public SourceControlPresenter
            (
                VBE vbe, 
                AddIn addin, 
                IConfigurationService<SourceControlConfiguration> configService,
                ISourceControlView view, 
                IChangesPresenter changesPresenter,
                IBranchesPresenter branchesPresenter,
                ISettingsPresenter settingsPresenter,
                IFolderBrowserFactory folderBrowserFactory,
                ISourceControlProviderFactory providerFactory
            ) 
            : base(vbe, addin, view)
        {
            _configService = configService;
            _config = _configService.LoadConfiguration();

            _changesPresenter = changesPresenter;
            
            _branchesPresenter = branchesPresenter;
            _settingsPresenter = settingsPresenter;
            _folderBrowserFactory = folderBrowserFactory;
            _providerFactory = providerFactory;
            _branchesPresenter.BranchChanged += _branchesPresenter_BranchChanged;

            _view = view;

            _view.RefreshData += OnRefreshChildren;
            _view.OpenWorkingDirectory += OnOpenWorkingDirectory;
            _view.InitializeNewRepository += OnInitNewRepository;
        }

        private void _branchesPresenter_BranchChanged(object sender, EventArgs e)
        {
            _changesPresenter.Refresh();
        }

        private void OnInitNewRepository(object sender, EventArgs e)
        {
            using (var folderPicker = _folderBrowserFactory.CreateFolderBrowser(RubberduckUI.SourceControl_CreateNewRepo, true, Environment.SpecialFolder.MyComputer))
            {
                if (folderPicker.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                var project = this.VBE.ActiveVBProject;

                _provider = _providerFactory.CreateProvider(project);
                var repo = _provider.InitVBAProject(folderPicker.SelectedPath);

                _provider = _providerFactory.CreateProvider(project, repo);

                AddRepoToConfig((Repository)repo);

                SetChildPresenterSourceControlProviders(_provider);
                _view.Status = RubberduckUI.Online;
            }
        }

        private void OnOpenWorkingDirectory(object sender, EventArgs e)
        {
            using (var folderPicker = _folderBrowserFactory.CreateFolderBrowser(RubberduckUI.SourceControl_OpenWorkingDirectory, false, Environment.SpecialFolder.MyComputer))
            {
                if (folderPicker.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                var project = this.VBE.ActiveVBProject;
                var repo = new Repository(project.Name, folderPicker.SelectedPath, string.Empty);
                _provider = _providerFactory.CreateProvider(project, repo);

                AddRepoToConfig(repo);

                SetChildPresenterSourceControlProviders(_provider);
                _view.Status = RubberduckUI.Online;
            }
        }

        private void AddRepoToConfig(Repository repo)
        {
            _config.Repositories.Add(repo);
            _configService.SaveConfiguration(_config, false);
        }

        private void OnRefreshChildren(object sender, EventArgs e)
        {
            RefreshChildren();
        }

        public void RefreshChildren()
        {
            if (!ValidRepoExists())
            {
                _view.Status = RubberduckUI.Offline;
                return;
            }

            try
            {
                _provider = _providerFactory.CreateProvider(this.VBE.ActiveVBProject,
                    _config.Repositories.First(repo => repo.Name == this.VBE.ActiveVBProject.Name));
            }
            catch (SourceControlException ex)
            {
                //todo: report failure to user and prompt to create or browse
                _provider = _providerFactory.CreateProvider(this.VBE.ActiveVBProject);
            }

            SetChildPresenterSourceControlProviders(_provider);

            _view.Status = RubberduckUI.Online;
        }

        private void SetChildPresenterSourceControlProviders(ISourceControlProvider provider)
        {
            _branchesPresenter.Provider = provider;
            _changesPresenter.Provider = provider;
            _settingsPresenter.Provider = provider;

            _branchesPresenter.RefreshView();
            _changesPresenter.Refresh();
            // Purposely not refreshing settingsPresenter.
            //  Settings it's provider doesn't affect it's view.
        }

        private bool ValidRepoExists()
        {
            if (_config.Repositories == null)
            {
                return false;
            }
            else
            {
                var possibleRepos = _config.Repositories.Where(repo => repo.Name == this.VBE.ActiveVBProject.Name);
                var possibleCount = possibleRepos.Count();

                if (possibleCount == 0 || possibleCount > 1)
                {
                    //todo: if none are found, prompt user to create one
                    //todo: more than one are found, prompt for correct one

                    return false;
                }
            }

            return true;
        }
    }
}
