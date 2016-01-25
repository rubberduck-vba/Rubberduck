using Ninject.Modules;

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

            Bind<ISourceControlView>().To<SourceControlPanel>().InSingletonScope();

            Bind<IFailedMessageView>().To<FailedActionControl>().InSingletonScope();
            Bind<ILoginView>().To<LoginControl>().InSingletonScope();

            Bind<IChangesView>().To<ChangesControl>().InSingletonScope();
            Bind<IUnsyncedCommitsView>().To<UnsyncedCommitsControl>().InSingletonScope();
            Bind<ISettingsView>().To<SettingsControl>().InSingletonScope();
            Bind<IBranchesView>().To<BranchesControl>().InSingletonScope();

            Bind<ICloneRepositoryView>().To<CloneRepositoryForm>();
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
