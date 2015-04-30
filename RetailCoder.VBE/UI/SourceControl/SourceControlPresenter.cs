using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Rubberduck.SourceControl;
using Rubberduck.Config;

namespace Rubberduck.UI.SourceControl
{
    public class SourceControlPresenter : DockablePresenterBase
    {
        private readonly IChangesPresenter _changesPresenter;
        private readonly IBranchesPresenter _branchesPresenter;
        private readonly ISourceControlView _view;
        private readonly IConfigurationService<SourceControlConfiguration> _configService;
        private SourceControlConfiguration _config;

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
            throw new NotImplementedException();
        }

        private void OnOpenWorkingDirectory(object sender, EventArgs e)
        {
            throw new NotImplementedException();
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

            ISourceControlProvider provider;

            try
            {
                provider = new GitProvider(this.VBE.ActiveVBProject, _config.Repositories.First());
            }
            catch (SourceControlException ex)
            {
                //todo: report failure to user and prompt to create or browse
                provider = new GitProvider(this.VBE.ActiveVBProject);
            }

            _branchesPresenter.Provider = provider;
            _changesPresenter.Provider = provider;

            _branchesPresenter.RefreshView();
            _changesPresenter.Refresh();

            _view.Status = "Online";
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
