using Castle.MicroKernel.Registration;
using Castle.Windsor;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Refactorings.DeleteDeclarations;
using Rubberduck.Refactorings;
using Castle.MicroKernel.SubSystems.Configuration;
using Castle.Facilities.TypedFactory;

namespace RubberduckTests.Refactoring.DeleteDeclarations
{
    public class DeleteDeclarationsTestsResolver : IWindsorInstaller
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IRewritingManager _rewritingManager;

        private IWindsorContainer _container;

        public DeleteDeclarationsTestsResolver(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;

            _rewritingManager = rewritingManager;

            Install(new WindsorContainer(), null);
        }


        public void Install(IWindsorContainer container, IConfigurationStore store)
            => Install(container);

        public T Resolve<T>() where T : class => _container.Resolve<T>() as T;

        private void Install(IWindsorContainer container)
        {
            _container = container;
            RegisterInstances(container);
            RegisterSingletonObjects(container);
            RegisterInterfaceToImplementationPairsSingleton(container);
        }

        private void RegisterInstances(IWindsorContainer container)
        {
            container.Kernel.Register(Component.For<IDeclarationFinderProvider, RubberduckParserState>().Instance(_declarationFinderProvider));
            container.Kernel.Register(Component.For<IRewritingManager>().Instance(_rewritingManager));
        }

        private static void RegisterSingletonObjects(IWindsorContainer container)
        {
            container.Kernel.Register(Component.For<DeleteDeclarationsRefactoringAction>());
            container.Kernel.Register(Component.For<DeleteUDTMembersRefactoringAction>());
            container.Kernel.Register(Component.For<DeleteEnumMembersRefactoringAction>());
            container.Kernel.Register(Component.For<DeleteModuleElementsRefactoringAction>());
            container.Kernel.Register(Component.For<DeleteProcedureScopeElementsRefactoringAction>());
        }

        private static void RegisterInterfaceToImplementationPairsSingleton(IWindsorContainer container)
        {
            container.Kernel.Register(Component.For<IDeclarationDeletionTargetFactory>()
                .ImplementedBy<DeclarationDeletionTargetFactory>());

            container.Kernel.Register(Component.For<IDeleteDeclarationEndOfStatementContentModifierFactory>()
                .ImplementedBy<DeleteDeclarationEndOfStatementContentModifierFactory>());

            container.Kernel.Register(Component.For<IEOSContextContentProviderFactory>()
                .ImplementedBy<EOSContextContentProviderFactory>());
        }
    }
}
