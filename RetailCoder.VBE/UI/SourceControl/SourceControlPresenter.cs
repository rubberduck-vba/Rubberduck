using System;
using System.Linq;
using System.Security;
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
        private readonly IUnsyncedCommitsPresenter _unsyncedPresenter;

        private readonly IFolderBrowserFactory _folderBrowserFactory;
        private readonly ISourceControlProviderFactory _providerFactory;
        private readonly ISourceControlView _view;
        private readonly IConfigurationService<SourceControlConfiguration> _configService;
        private readonly SourceControlConfiguration _config;

        private ISourceControlProvider _provider;
        private readonly IFailedMessageView _failedMessageView;
        private readonly ILoginView _loginView;

        public SourceControlPresenter(
            VBE vbe, AddIn addin, 
            IConfigurationService<SourceControlConfiguration> configService, 
            ISourceControlView view, 
            IChangesPresenter changesPresenter, 
            IBranchesPresenter branchesPresenter, 
            ISettingsPresenter settingsPresenter, 
            IUnsyncedCommitsPresenter unsyncedPresenter, 
            IFolderBrowserFactory folderBrowserFactory, 
            ISourceControlProviderFactory providerFactory, 
            IFailedMessageView failedMessageView, 
            ILoginView loginView
            )
            : base(vbe, addin, view)
        {
            _configService = configService;
            _config = _configService.LoadConfiguration();

            _changesPresenter = changesPresenter;
            _changesPresenter.ActionFailed += OnActionFailed;

            _branchesPresenter = branchesPresenter;
            _branchesPresenter.ActionFailed += OnActionFailed;

            _settingsPresenter = settingsPresenter;
            _settingsPresenter.ActionFailed += OnActionFailed;

            _unsyncedPresenter = unsyncedPresenter;
            _unsyncedPresenter.ActionFailed += OnActionFailed;

            _folderBrowserFactory = folderBrowserFactory;
            _providerFactory = providerFactory;
            _branchesPresenter.BranchChanged += _branchesPresenter_BranchChanged;

            _loginView = loginView;
            _loginView.Confirm += _loginView_Confirm;

            _failedMessageView = failedMessageView;
            _failedMessageView.DismissSecondaryPanel += DismissSecondaryPanel;

            _view = view;
            _view.SecondaryPanel = _failedMessageView;

            _view.RefreshData += OnRefreshChildren;
            _view.OpenWorkingDirectory += OnOpenWorkingDirectory;
            _view.InitializeNewRepository += OnInitNewRepository;
        }

        public override void Show()
        {
            RefreshChildren();
            base.Show();
        }

        private void _loginView_Confirm(object sender, EventArgs e)
        {
            var pwd = new SecureString();
            foreach (var c in _loginView.Password.ToCharArray())
            {
                pwd.AppendChar(c);
            }

            pwd.MakeReadOnly();

           var creds = new SecureCredentials(_loginView.Username, pwd);

           _provider = _providerFactory.CreateProvider(this.VBE.ActiveVBProject,
                    _config.Repositories.First(repo => repo.Name == this.VBE.ActiveVBProject.Name),
                    creds);

            SetChildPresenterSourceControlProviders(_provider);
        }

        private void DismissSecondaryPanel(object sender, EventArgs e)
        {
            _view.SecondaryPanelVisible = false;

            var panel = _view.SecondaryPanel as ILoginView;
            if (panel != null)
            {
                panel.Password = string.Empty;
            }
        }

        private void OnActionFailed(object sender, ActionFailedEventArgs e)
        {
            ShowSecondaryPanel(e.Title, e.Message);
        }

        private void _branchesPresenter_BranchChanged(object sender, EventArgs e)
        {
            _changesPresenter.RefreshView();
        }

        private void OnInitNewRepository(object sender, EventArgs e)
        {
            using (var folderPicker = _folderBrowserFactory.CreateFolderBrowser((RubberduckUI.SourceControl_CreateNewRepo)))
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
            using (var folderPicker = _folderBrowserFactory.CreateFolderBrowser(RubberduckUI.SourceControl_OpenWorkingDirectory, false))
            {
                if (folderPicker.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                var project = this.VBE.ActiveVBProject;
                var repo = new Repository(project.Name, folderPicker.SelectedPath, string.Empty);

                try
                {
                    _provider = _providerFactory.CreateProvider(project, repo);
                }
                catch (SourceControlException ex)
                {
                    ShowSecondaryPanel(ex.Message, ex.InnerException.Message);
                    return;
                }

                AddRepoToConfig(repo);

                SetChildPresenterSourceControlProviders(_provider);
                _view.Status = RubberduckUI.Online;
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

        private void ShowSecondaryPanel(string title, string message)
        {
            //this comes out of LibGit2Sharp it would be difficult to check this a different way
            const string unauthorizedMessage = "Request failed with status code: 401";

            if (message == unauthorizedMessage)
            {
                _failedMessageView.DismissSecondaryPanel -= DismissSecondaryPanel;
                _loginView.DismissSecondaryPanel += DismissSecondaryPanel;

                _view.SecondaryPanel = _loginView;
            }
            else
            {
                _loginView.DismissSecondaryPanel -= DismissSecondaryPanel;
                _failedMessageView.DismissSecondaryPanel += DismissSecondaryPanel;

                _failedMessageView.Message = string.Format("{0}{1}{2}", title, Environment.NewLine, message);

                _view.SecondaryPanel = _failedMessageView;
            }

            _view.SecondaryPanelVisible = true;
        }

        private void SetChildPresenterSourceControlProviders(ISourceControlProvider provider)
        {
            _branchesPresenter.Provider = provider;
            _changesPresenter.Provider = provider;
            _settingsPresenter.Provider = provider;
            _unsyncedPresenter.Provider = provider;

            _branchesPresenter.RefreshView();
            _changesPresenter.RefreshView();
            _unsyncedPresenter.RefreshView();
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
