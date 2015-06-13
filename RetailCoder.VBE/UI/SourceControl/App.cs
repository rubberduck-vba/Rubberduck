using NetOffice.VBIDEApi;
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
            var settingsPresenter = new SettingsPresenter(settingsView, configService, new DialogFactory());

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
                    new DialogFactory(), 
                    new SourceControlProviderFactory()
                );
        }

        public void ShowWindow()
        {
            _sourceControlPresenter.RefreshChildren();
            _sourceControlPresenter.Show();
        }
    }
}
