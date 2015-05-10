using System;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Config;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public class SourceControlPresenter : DockablePresenterBase
    {
        private readonly IChangesPresenter _changesPresenter;
        private readonly IBranchesPresenter _branchesPresenter;
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
                IBranchesPresenter branchesPresenter           
            ) 
            : base(vbe, addin, view)
        {
            _configService = configService;
            _config = _configService.LoadConfiguration();

            _changesPresenter = changesPresenter;
            _branchesPresenter = branchesPresenter;
            _view = view;

            _view.RefreshData += OnRefreshChildren;
            _view.OpenWorkingDirectory += OnOpenWorkingDirectory;
            _view.InitializeNewRepository += OnInitNewRepository;
        }

        private void OnInitNewRepository(object sender, EventArgs e)
        {
            using (var folderPicker = new FolderBrowserDialog())
            {
                folderPicker.Description = "Create New Repository";
                folderPicker.RootFolder = Environment.SpecialFolder.MyDocuments;
                folderPicker.ShowNewFolderButton = true;

                if (folderPicker.ShowDialog() == DialogResult.OK)
                {
                    var project = this.VBE.ActiveVBProject;

                    _provider = new GitProvider(project);
                    var repo = _provider.InitVBAProject(folderPicker.SelectedPath);

                    _provider = new GitProvider(project, repo);

                    AddRepoToConfig((Repository)repo);

                    SetChildPresenterSourceControlProviders(_provider);
                }
            }
        }

        private void OnOpenWorkingDirectory(object sender, EventArgs e)
        {
            using (var folderPicker = new FolderBrowserDialog())
            {
                folderPicker.Description = "Open Working Directory";
                folderPicker.RootFolder = Environment.SpecialFolder.MyDocuments;
                folderPicker.ShowNewFolderButton = false;

                if (folderPicker.ShowDialog() == DialogResult.OK)
                {
                    var project = this.VBE.ActiveVBProject;
                    var repo = new Repository(project.Name, folderPicker.SelectedPath, string.Empty);
                    _provider = new GitProvider(project, repo);

                    AddRepoToConfig(repo);

                    SetChildPresenterSourceControlProviders(_provider);
                }
            }
        }

        private void AddRepoToConfig(Repository repo)
        {
            _config.Repositories.Add(repo);
            _configService.SaveConfiguration(_config);
        }

        private void OnRefreshChildren(object sender, EventArgs e)
        {
            RefreshChildren();
        }

        public void RefreshChildren()
        {
            if (!ValidRepoExists())
            {
                _view.Status = "Offline";
                return;
            }

            try
            {
                _provider = new GitProvider(this.VBE.ActiveVBProject, _config.Repositories.First());
            }
            catch (SourceControlException ex)
            {
                //todo: report failure to user and prompt to create or browse
                _provider = new GitProvider(this.VBE.ActiveVBProject);
            }

            SetChildPresenterSourceControlProviders(_provider);

            _view.Status = "Online";
        }

        private void SetChildPresenterSourceControlProviders(ISourceControlProvider provider)
        {
            _branchesPresenter.Provider = provider;
            _changesPresenter.Provider = provider;

            _branchesPresenter.RefreshView();
            _changesPresenter.Refresh();
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
