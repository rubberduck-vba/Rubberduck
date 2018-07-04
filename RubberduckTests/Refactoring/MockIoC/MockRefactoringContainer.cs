using System;
using System.Linq;
using Castle.Facilities.TypedFactory;
using Castle.MicroKernel.ModelBuilder.Inspectors;
using Castle.MicroKernel.Registration;
using Castle.MicroKernel.Resolvers.SpecializedResolvers;
using Castle.MicroKernel.SubSystems.Configuration;
using Castle.Windsor;
using Moq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.Symbols.DeclarationLoaders;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.UI.Refactorings;

namespace RubberduckTests.Refactoring.MockIoC
{
    internal class RefactoringContainerInstaller : IWindsorInstaller
    {
        internal static IWindsorContainer GetContainer()
        {
            return new WindsorContainer().Install(new RefactoringContainerInstaller());
        }

        public void Install(IWindsorContainer container, IConfigurationStore store)
        {
            SetUpCollectionResolver(container);
            ActivateAutoMagicFactories(container);
            DeactivatePropertyInjection(container);

            container.Kernel.Resolver.AddSubResolver(
                new AutoMoqResolver(
                    container.Kernel));
            container.Register(Component
                .For(typeof(Mock<>)));

            container.Register(Component.For<RubberduckParserState, IParseTreeProvider, IDeclarationFinderProvider>()
                .ImplementedBy<RubberduckParserState>()
                .LifestyleSingleton());
            container.Register(Component.For<DeclarationFinder>()
                .ImplementedBy<ConcurrentlyConstructedDeclarationFinder>()
                .LifestyleTransient());
            RegisterParsingEngine(container);

            container.Register(Component.For<IRefactoringPresenterFactory>().AsFactory().LifestyleTransient());
            container.Register(Component.For<IRefactoringDialogFactory>().AsFactory().LifestyleTransient());

            container.Register(Classes
                .FromAssemblyContaining(typeof(IParseCoordinator))
                .InSameNamespaceAs(typeof(IParseCoordinator), true)
                .WithService.DefaultInterfaces()
                );
            container.Register(Classes
                .FromAssemblyContaining<IRefactoring>()
                .IncludeNonPublicTypes()
                .InSameNamespaceAs(typeof(IRefactoring), true)
                .WithService.DefaultInterfaces()
                .LifestyleTransient());

            container.Register(Classes
                .FromAssemblyContaining<RefactoringDialogBase>()
                .IncludeNonPublicTypes()
                .InSameNamespaceAs(typeof(RefactoringDialogBase), true)
                .WithService.DefaultInterfaces()
                .LifestyleTransient());

            /*
            container.Register(Component.For<IRefactoringView>().ImplementedBy(typeof(RenameView)));
            container.Register(Component.For(typeof(IRefactoringViewModel<>))
                .ImplementedBy(typeof(RenameViewModel)).LifestyleTransient());
            container.Register(Component.For(typeof(IRefactoringDialog<,,>)).ImplementedBy(typeof(RenameDialog))
                .LifestyleTransient());
            container.Register(Component.For(typeof(IRefactoringPresenter<,,,>)).ImplementedBy(typeof(RenamePresenter))
                .LifestyleTransient());
            */
        }

        private void SetUpCollectionResolver(IWindsorContainer container)
        {
            container.Kernel.Resolver.AddSubResolver(new CollectionResolver(container.Kernel, true));
        }

        private void ActivateAutoMagicFactories(IWindsorContainer container)
        {
            container.Kernel.AddFacility<TypedFactoryFacility>();
        }

        private void DeactivatePropertyInjection(IWindsorContainer container)
        {
            // We don't want to inject properties, only ctors. 
            //There are too many properties that would be injected otherwise, which causes code to execute at resolve time.
            var propInjector = container.Kernel.ComponentModelBuilder
                .Contributors
                .OfType<PropertiesDependenciesModelInspector>()
                .Single();
            container.Kernel.ComponentModelBuilder.RemoveContributor(propInjector);
        }

        private void RegisterParsingEngine(IWindsorContainer container)
        {
            RegisterCustomDeclarationLoadersToParser(container);

            container.Register(Component.For<ICOMReferenceSynchronizer, IProjectReferencesProvider>()
                .ImplementedBy<COMReferenceSynchronizer>()
                .DependsOn(Dependency.OnValue<string>(null))
                .LifestyleSingleton());
            container.Register(Component.For<IBuiltInDeclarationLoader>()
                .ImplementedBy<BuiltInDeclarationLoader>()
                .LifestyleSingleton());
            container.Register(Component.For<IDeclarationResolveRunner>()
                .ImplementedBy<DeclarationResolveRunner>()
                .LifestyleSingleton());
            container.Register(Component.For<IModuleToModuleReferenceManager>()
                .ImplementedBy<ModuleToModuleReferenceManager>()
                .LifestyleSingleton());
            container.Register(Component.For<ISupertypeClearer>()
                .ImplementedBy<SupertypeClearer>()
                .LifestyleSingleton());
            container.Register(Component.For<IParserStateManager>()
                .ImplementedBy<ParserStateManager>()
                .LifestyleSingleton());
            container.Register(Component.For<IParseRunner>()
                .ImplementedBy<ParseRunner>()
                .LifestyleSingleton());
            container.Register(Component.For<IParsingStageService>()
                .ImplementedBy<ParsingStageService>()
                .LifestyleSingleton());
            container.Register(Component.For<IParsingCacheService>()
                .ImplementedBy<ParsingCacheService>()
                .LifestyleSingleton());
            container.Register(Component.For<IProjectManager>()
                .ImplementedBy<RepositoryProjectManager>()
                .LifestyleSingleton());
            container.Register(Component.For<IReferenceRemover>()
                .ImplementedBy<ReferenceRemover>()
                .LifestyleSingleton());
            container.Register(Component.For<IReferenceResolveRunner>()
                .ImplementedBy<ReferenceResolveRunner>()
                .LifestyleSingleton());
            container.Register(Component.For<IParseCoordinator>()
                .ImplementedBy<ParseCoordinator>()
                .LifestyleSingleton());

            container.Register(Component.For<Func<IVBAPreprocessor>>()
                .Instance(() => new VBAPreprocessor(7.0)));
        }
        private void RegisterCustomDeclarationLoadersToParser(IWindsorContainer container)
        {
            container.Register(Classes.FromAssemblyContaining<ICustomDeclarationLoader>()
                .BasedOn<ICustomDeclarationLoader>()
                .WithService.Base()
                .LifestyleSingleton());
        }
    }
}
