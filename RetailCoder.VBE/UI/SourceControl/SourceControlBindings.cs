
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            Bind<IConfigurationService<SourceControlConfiguration>>().To<SourceControlConfigurationService>();

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

            //presenters
            Bind<IChangesPresenter>().To<ChangesPresenter>();
            Bind<IBranchesPresenter>().To<BranchesPresenter>();
            Bind<ISettingsPresenter>().To<SettingsPresenter>();
            Bind<IUnsyncedCommitsPresenter>().To<UnsyncedCommitsPresenter>();

            //factories
            Bind<ISourceControlProviderFactory>().To<SourceControlProviderFactory>();
            Bind<IFolderBrowserFactory>().To<DialogFactory>();
        }
    }
}
