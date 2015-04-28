using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.SourceControl;
using Rubberduck.UI.SourceControl;
using Rubberduck.Config;
using Microsoft.Vbe.Interop;

namespace Rubberduck.SourceControl
{
    class App
    {
        private SourceControlPresenter _sourceControlPresenter;
        private ISourceControlView _sourceControlView;

        internal App(
                    VBE vbe, 
                    AddIn addIn, 
                    IConfigurationService configService, 
                    IChangesView changesView, 
                    IUnSyncedCommitsView unsyncedCommitsView, 
                    ISettingsView settingsView,
                    IBranchesView branchesView, 
                    ICreateBranchView createBranchView,
                    IMergeView mergeView
                )
        {
             _sourceControlView = new SourceControlPanel(branchesView, changesView, unsyncedCommitsView, settingsView);
             
            var repo = new Repository
            (
                "SourceControlTest", 
                @"C:\Users\Christopher\Documents\SourceControlTest",
                @"https://github.com/ckuhn203/SourceControlTest.git"
            );
            var gitProvider = new GitProvider(vbe.ActiveVBProject, repo);
            var changesPresenter = new ChangesPresenter(gitProvider, changesView);
            var branchesPresenter = new BranchesPresenter(gitProvider, branchesView, createBranchView, mergeView);
            branchesPresenter.RefreshView();

            _sourceControlPresenter = new SourceControlPresenter(vbe, addIn, _sourceControlView, changesPresenter, branchesPresenter);
        }

        public void ShowWindow()
        {
            _sourceControlPresenter.RefreshChildren();
            _sourceControlPresenter.Show();
        }
    }
}
