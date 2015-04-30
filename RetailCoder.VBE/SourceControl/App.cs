using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.SourceControl;
using Rubberduck.UI.SourceControl;
using Rubberduck.Config;
using Microsoft.Vbe.Interop;
using Rubberduck.UI;

namespace Rubberduck.SourceControl
{
    class App
    {
        private SourceControlPresenter _sourceControlPresenter;
        private ISourceControlView _sourceControlView;

        internal App(
                    VBE vbe, 
                    AddIn addIn, 
                    IConfigurationService<SourceControlConfiguration> configService, 
                    IChangesView changesView, 
                    IUnSyncedCommitsView unsyncedCommitsView, 
                    ISettingsView settingsView,
                    IBranchesView branchesView, 
                    ICreateBranchView createBranchView,
                    IMergeView mergeView
                )
        {
             _sourceControlView = new SourceControlPanel(branchesView, changesView, unsyncedCommitsView, settingsView);

            var changesPresenter = new ChangesPresenter(changesView);
            var branchesPresenter = new BranchesPresenter(branchesView, createBranchView, mergeView);

            _sourceControlPresenter = new SourceControlPresenter(vbe, addIn, configService, new FolderBrowserDialog(), _sourceControlView, changesPresenter, branchesPresenter);
        }

        public void ShowWindow()
        {
            _sourceControlPresenter.RefreshChildren();
            _sourceControlPresenter.Show();
        }
    }
}
