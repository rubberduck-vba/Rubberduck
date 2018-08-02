using System;
using System.Collections;
using System.Linq;
using System.Reflection;
using Castle.Core;
using Castle.Facilities.TypedFactory;
using Castle.MicroKernel;
using Castle.MicroKernel.ComponentActivator;
using Castle.MicroKernel.Context;
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
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Refactorings;
using Rubberduck.UI.Refactorings.Rename;

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

            container.Register(Component.For<IRefactoringPresenterFactory>()
                .AsFactory()
                .LifestyleTransient()
                );

            container.Register(Component.For<IRefactoringDialogFactory>()
                .AsFactory()
                .LifestyleTransient()
                );

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

            container.Register(Component
                .For(typeof(IRefactoringDialog<,,>))
                .ImplementedBy(typeof(RefactoringDialogStub<,,>))
                .LifestyleSingleton()
            );

            container.Register(Component
                .For(typeof(IRefactoringView<>))
                .ImplementedBy(typeof(RefactoringViewStub<>))
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
                .LifestyleTransient());
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

    /*
    public class MockActivator : DefaultComponentActivator
    {
        public MockActivator(ComponentModel model, IKernelInternal kernel, ComponentInstanceDelegate onCreation,
            ComponentInstanceDelegate onDestruction) : base(model, kernel, onCreation, onDestruction) { }

        public override object Create(CreationContext context, Burden burden)
        {
            if (context.RequestedType.IsGenericType && context.RequestedType.GetGenericTypeDefinition() == typeof(IRefactoringView<>))
            {
                var refactoringModel = context.AdditionalArguments["model"];
                var modelType = refactoringModel.GetType();

                var viewType = typeof(IRefactoringView<>).MakeGenericType(modelType);
                var mockViewType = typeof(Mock<>).MakeGenericType(viewType);
                var mockView = (Mock)Kernel.Resolve(mockViewType);

                return mockView.Object;
            }
            else if (context.RequestedType.IsGenericType && context.RequestedType.GetGenericTypeDefinition() == typeof(IRefactoringDialog<,,>))
            {
                var refactoringModel = context.AdditionalArguments["model"];
                var modelType = refactoringModel.GetType();

                var viewModelType = context.RequestedType.GenericTypeArguments[2];
                var viewModel = Kernel.Resolve(viewModelType, new Arguments {{"model", refactoringModel}});

                var viewType = typeof(IRefactoringView<>).MakeGenericType(modelType);
                var mockViewType = typeof(Mock<>).MakeGenericType(viewType);
                var mockView = (Mock)Kernel.Resolve(mockViewType);

                var dialogType = typeof(IRefactoringDialog<,,>).MakeGenericType(modelType, viewType, viewModelType);
                var mockDialogType = typeof(Mock<>).MakeGenericType(dialogType);

                var args = new Arguments(new
                {
                    model = refactoringModel,
                    view = mockView,
                    viewModel = viewModel
                });
                var mockDialog = (Mock<IRefactoringDialog<RenameModel, IRefactoringView<RenameModel>, RenameViewModel>>)Kernel.Resolve(mockDialogType);

                mockDialog.SetupAllProperties();
                return mockDialog.Object;
                
                //Kernel.Resolve(typeof(IRefactoringDialog<,,>).MakeGenericType(modelType, viewType, viewModelType),
                //    new Arguments {{"model", refactoringModel}, {"view", mockView.Object}, {"viewModel", viewModel}});
                /*
                var modelType = context.RequestedType.GenericTypeArguments[0];
                var viewModelType = context.RequestedType.GenericTypeArguments[2];

                var viewType = typeof(IRefactoringView<>).MakeGenericType(modelType);
                var mockViewType = typeof(Mock<>).MakeGenericType(viewType);
                var mockView = Kernel.Resolve(mockViewType);

                var requestedGeneric = context.RequestedType.GetGenericTypeDefinition();
                var returnType = requestedGeneric.MakeGenericType(modelType, mockViewType, viewModelType);

                Activator.CreateInstance(returnType);

                //var viewModelType = typeof(IRefactoringViewModel<>).MakeGenericType(modelType);
                //var mockViewModelType = typeof(Mock<>).MakeGenericType(viewModelType);
                //var mock = Kernel.Resolve(mockViewModelType);
                * /
            }

            return base.Create(context, burden);
        }
    }
    */
}
