using Microsoft.Vbe.Interop;
using Rubberduck.Settings;

namespace Rubberduck.UI.SourceControl
{
    class App
    {
        private readonly SourceControlPresenter _sourceControlPresenter;
        private ISourceControlView _sourceControlView;

        internal App(
                    VBE vbe, 
                    AddIn addIn, 
                    IConfigurationService<SourceControlConfiguration> configService, 
                    IChangesView changesView, 
                    IUnsyncedCommitsView unsyncedCommitsView, 
                    ISettingsView settingsView,
                    IBranchesView branchesView, 
                    ICreateBranchView createBranchView,
                    IDeleteBranchView deleteBranchView,
                    IMergeView mergeView
                )
        {
            var failedActionView = new FailedActionControl();
       
             _sourceControlView = new SourceControlPanel(branchesView, changesView, unsyncedCommitsView, settingsView, failedActionView);
            var changesPresenter = new ChangesPresenter(changesView);
            var branchesPresenter = new BranchesPresenter(branchesView, createBranchView, deleteBranchView, mergeView);
            var settingsPresenter = new SettingsPresenter(settingsView, configService, new DialogFactory());
            var unsyncedPresenter = new UnsyncedCommitsPresenter(unsyncedCommitsView);

            _sourceControlPresenter = 
                new SourceControlPresenter
                (
                    vbe, 
                    addIn, 
                    configService, 
                    _sourceControlView, 
                    changesPresenter, 
                    branchesPresenter, 
                    settingsPresenter, 
                    unsyncedPresenter,
                    new DialogFactory(), 
                    new SourceControlProviderFactory(),
                    failedActionView
                );
        }

        public void ShowWindow()
        {
            _sourceControlPresenter.RefreshChildren();
            _sourceControlPresenter.Show();
        }
    }
}
