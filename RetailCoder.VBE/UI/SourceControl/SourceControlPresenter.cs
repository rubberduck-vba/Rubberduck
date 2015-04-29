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
        }

        private void OnRefreshChildren(object sender, EventArgs e)
        {
            RefreshChildren();
        }

        public void RefreshChildren()
        {
            //todo: get repo from config for the active project
            //todo: send a provider down into the child presenters

            _branchesPresenter.RefreshView();
            _changesPresenter.Refresh();
        }
    }
}
