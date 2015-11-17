
using Ninject;
using Ninject.Modules;
using Rubberduck.Settings;

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
            // todo: check on note below
            // ninject is complaining about also having a SourceControlProviderFactoryProxy and a FolderBrowserFactoryProxy
            // I'm unsure about commenting these out. I have a feeling that it's not the "right thing", but everything seems to work.

            //Bind<ISourceControlProviderFactory>().To<SourceControlProviderFactory>();
            //Bind<IFolderBrowserFactory>().To<DialogFactory>();
        }
    }
}
