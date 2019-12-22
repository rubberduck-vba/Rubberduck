using System.Globalization;
using System.Linq;
using System.Threading;
using Castle.Facilities.TypedFactory;
using Castle.MicroKernel.ModelBuilder.Inspectors;
using Castle.MicroKernel.Registration;
using Castle.MicroKernel.Resolvers.SpecializedResolvers;
using Castle.MicroKernel.SubSystems.Configuration;
using Castle.Windsor;
using Moq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.Symbols.DeclarationLoaders;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.ComReferenceLoading;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA.DeclarationResolving;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Parsing.VBA.Parsing.ParsingExceptions;
using Rubberduck.Parsing.VBA.ReferenceManagement;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Refactorings;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace RubberduckTests.Refactoring.MockIoC
{
    internal class RefactoringContainerInstaller : IWindsorInstaller
    {
        /// <summary>
        /// ThreadLocal to enable separate resolution and disjoint object graphs for separate threads.
        /// This should avoid concurrency issues when running tests in parallel.
        /// </summary>
        private static ThreadLocal<IWindsorContainer> _container = new ThreadLocal<IWindsorContainer>(() => new WindsorContainer().Install(new RefactoringContainerInstaller()));
        internal static IWindsorContainer GetContainer()
        {
            return _container.Value;
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
                .For(typeof(Mock<>))
                .LifestyleSingleton());

            container.Register(Component.For<RubberduckParserState, IParseTreeProvider, IDeclarationFinderProvider>()
                .ImplementedBy<RubberduckParserState>()
                .LifestyleSingleton());
            container.Register(Component.For<DeclarationFinder>()
                .ImplementedBy<ConcurrentlyConstructedDeclarationFinder>()
                .LifestyleTransient());
            RegisterParsingEngine(container);

            container.Register(Component.For<IRefactoringPresenterFactory>()
                .AsFactory(f => f.SelectedWith(new AutoMoqFactorySelector()))
                .LifestyleSingleton()
                );

            container.Register(Component.For<IRefactoringDialogFactory>()
                .AsFactory(f => f.SelectedWith(new AutoMoqFactorySelector()))
                .LifestyleSingleton()
                );

            container.Register(Classes
                .FromAssemblyContaining(typeof(IParseCoordinator))
                .InSameNamespaceAs(typeof(IParseCoordinator), true)
                .WithService.DefaultInterfaces()
                );

            container.Register(Component
                .For(typeof(IRefactoringDialog<,,>))
                .ImplementedBy(typeof(RefactoringDialogStub<,,>))
                .LifestyleSingleton()
            );

            /*
            container.Register(Component
                .For(typeof(IRefactoringView<>))
                .ImplementedBy(typeof(RefactoringViewStub<>))
                .LifestyleSingleton()
            );
            */

            container.Register(Component
                .For(typeof(RenameViewModel), typeof(IRefactoringViewModel<RenameModel>))
                .ImplementedBy<RenameViewModel>()
                .LifestyleSingleton()
            );

            container.Register(Classes
                .FromAssemblyContaining<RefactoringDialogBase>()
                .IncludeNonPublicTypes()
                .InSameNamespaceAs(typeof(RefactoringDialogBase), true)
                .If(type => 
                    !type.Name.EndsWith("Dialog") 
                    && !type.Name.EndsWith("DialogBase")
                    && !type.Name.EndsWith("View")
                    && !type.Name.EndsWith("ViewBase"))
                .Unless(t => t.IsAbstract)
                .WithService.DefaultInterfaces()
                .LifestyleSingleton());
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

            container.Register(Component.For<ICompilationArgumentsProvider, ICompilationArgumentsCache>()
                .ImplementedBy<CompilationArgumentsCache>()
                .DependsOn(Dependency.OnComponent<ICompilationArgumentsProvider, CompilationArgumentsProvider>())
                .LifestyleSingleton());
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
            container.Register(Component.For<ITokenStreamPreprocessor>()
                .ImplementedBy<VBAPreprocessor>()
                .DependsOn(Dependency.OnComponent<ITokenStreamParser, VBAPreprocessorParser>())
                .LifestyleSingleton());
            container.Register(Component.For<VBAPredefinedCompilationConstants>()
                .ImplementedBy<VBAPredefinedCompilationConstants>()
                .DependsOn(Dependency.OnValue<double>(double.Parse("7.1", CultureInfo.InvariantCulture)))
                .LifestyleSingleton());
            container.Register(Component.For<VBAPreprocessorParser>()
                .ImplementedBy<VBAPreprocessorParser>()
                .DependsOn(Dependency.OnComponent<IParsePassErrorListenerFactory, PreprocessingParseErrorListenerFactory>())
                .LifestyleSingleton());
            container.Register(Component.For<ICommonTokenStreamProvider>()
                .ImplementedBy<SimpleVBAModuleTokenStreamProvider>()
                .LifestyleSingleton());
            container.Register(Component.For<IStringParser>()
                .ImplementedBy<TokenStreamParserStringParserAdapterWithPreprocessing>()
                .LifestyleSingleton());
            container.Register(Component.For<IModuleParser>()
                .ImplementedBy<ModuleParser>()
                .DependsOn(Dependency.OnComponent("codePaneSourceCodeProvider", typeof(CodePaneHandler)),
                    Dependency.OnComponent("attributesSourceCodeProvider", typeof(SourceFileHandlerComponentSourceCodeHandlerAdapter)))
                .LifestyleSingleton());
            container.Register(Component.For<ITypeLibWrapperProvider>()
                .ImplementedBy<TypeLibWrapperProvider>()
                .LifestyleSingleton());
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
