
using Ninject;
using Ninject.Modules;
using Rubberduck.Settings;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    class SourceControlBindings : NinjectModule
    {
        /// <summary>
        /// Loads the module into the kernel.
        /// </summary>
        public override void Load()
        {
            //ConfigurationService and Presenters are bound by convention

            //user controls (views)

            Bind<ISourceControlView>().To<SourceControlPanel>();

            Bind<IFailedMessageView>().To<FailedActionControl>();
            Bind<ILoginView>().To<LoginControl>();

            Bind<IChangesView>().To<ChangesControl>();
            Bind<IUnsyncedCommitsView>().To<UnsyncedCommitsControl>();
            Bind<ISettingsView>().To<SettingsControl>();
            Bind<IBranchesView>().To<BranchesControl>();

            Bind<ICreateBranchView>().To<CreateBranchForm>();
            Bind<IDeleteBranchView>().To<DeleteBranchForm>();
            Bind<IMergeView>().To<MergeForm>();

            //factories 

            // note: RubberduckModule sets up factory proxies by convention. 
            // Replace these factory proxies with our existing concrete implementations.
            Rebind<ISourceControlProviderFactory>().To<SourceControlProviderFactory>();
            Rebind<IFolderBrowserFactory>().To<DialogFactory>();
        }
    }
}
