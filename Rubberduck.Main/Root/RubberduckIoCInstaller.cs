using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Castle.Facilities.TypedFactory;
using Castle.MicroKernel.ModelBuilder.Inspectors;
using Castle.MicroKernel.Registration;
using Component = Castle.MicroKernel.Registration.Component;
using Castle.MicroKernel.Resolvers.SpecializedResolvers;
using Castle.MicroKernel.SubSystems.Configuration;
using Castle.Windsor;
using Rubberduck.AutoComplete;
using Rubberduck.CodeAnalysis.CodeMetrics;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation;
using Rubberduck.CodeAnalysis.Inspections.Logistics;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.ComClientLibrary.UnitTesting;
using Rubberduck.Common;
using Rubberduck.Common.Hotkeys;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.Symbols.DeclarationLoaders;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.ComReferenceLoading;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA.DeclarationResolving;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Parsing.VBA.Parsing.ParsingExceptions;
using Rubberduck.Parsing.VBA.ReferenceManagement;
using Rubberduck.Refactorings;
using Rubberduck.Runtime;
using Rubberduck.Settings;
using GeneralSettings = Rubberduck.Settings.GeneralSettings;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;
using IndenterSettings = Rubberduck.SmartIndenter.IndenterSettings;
using Rubberduck.UI;
using Rubberduck.UI.AddRemoveReferences;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.Command.MenuItems.CommandBars;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Controls;
using Rubberduck.UI.Settings;
using Rubberduck.UI.UnitTesting;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.Utility;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;
using Rubberduck.VBEditor.VbeRuntime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.UI.Refactorings.AnnotateDeclaration;

namespace Rubberduck.Root
{
    public class RubberduckIoCInstaller : IWindsorInstaller
    {
        private readonly IVBE _vbe;
        private readonly IAddIn _addin;
        private readonly GeneralSettings _initialSettings;
        private readonly IVbeNativeApi _vbeNativeApi;
        private readonly IBeepInterceptor _beepInterceptor;

        public RubberduckIoCInstaller(IVBE vbe, IAddIn addin, GeneralSettings initialSettings, IVbeNativeApi vbeNativeApi, IBeepInterceptor beepInterceptor)
        {
            _vbe = vbe;
            _addin = addin;
            _initialSettings = initialSettings;
            _vbeNativeApi = vbeNativeApi;
            _beepInterceptor = beepInterceptor;
        }

        //Guidelines and words of caution:

        //1) Please always specify the Lifestyle. The default is singleton, which can be confusing.
        //2) Before adding conventions, please read the Castle Windsor online documentation; there are a few gotchas.
        //3) The first binding wins; all further bindings are only used in multibinding, unless named.
        //4) The standard name of a binding is the full type name of the implementing class.


        public void Install(IWindsorContainer container, IConfigurationStore store)
        {
            SetUpCollectionResolver(container);
            ActivateAutoMagicFactories(container);
            OverridePropertyInjection(container);

            RegisterInstances(container);
            RegisterAppWithSpecialDependencies(container);
            RegisterUnitTestingComSide(container);

            container.Register(Component.For<Version>()
                     .UsingFactoryMethod(() => Assembly.GetExecutingAssembly().GetName().Version)
                     .LifestyleSingleton());
            container.Register(Component.For<IProjectsProvider, IProjectsRepository>()
                .ImplementedBy<ProjectsRepository>()
                .LifestyleSingleton());
            container.Register(Component.For<RubberduckParserState, IParseTreeProvider, IDeclarationFinderProvider, IParseManager, IParserStatusProvider>()
                .ImplementedBy<RubberduckParserState>()
                .LifestyleSingleton());
            container.Register(Component.For<ISelectionChangeService>()
                .ImplementedBy<SelectionChangeService>()
                .LifestyleSingleton());
            container.Register(Component.For<ISelectionService, ISelectionProvider>()
                .ImplementedBy<SelectionService>()
                .LifestyleSingleton());
            container.Register(Component.For<AutoCompleteService>()
                .LifestyleSingleton());
            container.Register(Component.For<IOperatingSystem>()
                .ImplementedBy<WindowsOperatingSystem>()
                .LifestyleSingleton());

            container.Register(Component.For<DeclarationFinder>()
                .ImplementedBy<ConcurrentlyConstructedDeclarationFinder>()
                .LifestyleTransient());

            RegisterSmartIndenter(container);
            RegisterSourceCodeHandlers(container);
            RegisterParsingEngine(container);
            RegisterTypeLibApi(container);

            container.Register(Component.For<ISelectedDeclarationProvider>()
                .ImplementedBy<SelectedDeclarationProvider>()
                .LifestyleSingleton());

            container.Register(Component.For<IRewritingManager>()
                .ImplementedBy<RewritingManager>()
                .LifestyleSingleton());
            container.Register(Component.For<IMemberAttributeRecovererWithSettableRewritingManager>()
                .ImplementedBy<MemberAttributeRecoverer>()
                .LifestyleSingleton());
            container.Register(Component.For<IAddComponentService>()
                .ImplementedBy<AddComponentService>()
                .DependsOn(Dependency.OnComponent("codePaneComponentSourceCodeProvider", typeof(CodeModuleComponentSourceCodeHandler)),
                    Dependency.OnComponent("attributesComponentSourceCodeProvider", typeof(SourceFileHandlerComponentSourceCodeHandlerAdapter)))
                .LifestyleSingleton());

            container.Register(Component.For<TestExplorerModel>()
                .LifestyleSingleton());
            container.Register(Component.For<IVBEInteraction>()
                .ImplementedBy<VBEInteraction>()
                .LifestyleSingleton());

            RegisterSettingsViewModel(container);
            RegisterRefactoringPreviewProviders(container);
            RegisterRefactoringDialogs(container);

            container.Register(Component.For<ISearchResultsWindowViewModel>()
                .ImplementedBy<SearchResultsWindowViewModel>()
                .LifestyleSingleton());
            container.Register(Component.For<SearchResultPresenterInstanceManager>()
                .LifestyleSingleton());
            
            RegisterDockablePresenters(container);
            RegisterDockableUserControls(container);

            RegisterCommands(container);
            RegisterCommandMenuItems(container);
            RegisterParentMenus(container);

            RegisterRubberduckCommandBar(container);
            RegisterRubberduckMenu(container);
            RegisterCodePaneContextMenu(container);
            RegisterFormDesignerContextMenu(container);
            RegisterFormDesignerControlContextMenu(container);
            RegisterProjectExplorerContextMenu(container);

            RegisterWindowsHooks(container);

            container.Register(Component.For<HotkeyFactory>()
                .LifestyleSingleton());
            container.Register(Component.For<ITestEngine>()
                .ImplementedBy<TestEngine>()
                .LifestyleSingleton());

            var assembliesToRegister = AssembliesToRegister().ToArray();

            RegisterConfiguration(container, assembliesToRegister);

            RegisterRequiredBinaryExtractors(container, assembliesToRegister);

            RegisterParseTreeInspections(container, assembliesToRegister);
            RegisterInspections(container, assembliesToRegister);
            RegisterQuickFixes(container, assembliesToRegister);
            RegisterRefactorings(container, assembliesToRegister);
            RegisterAutoCompletes(container, assembliesToRegister);
            RegisterCodeMetrics(container, assembliesToRegister);

            RegisterSpecialFactories(container);
            RegisterFactories(container, assembliesToRegister);

            ApplyDefaultInterfaceConvention(container, assembliesToRegister);
        }

        private static void RegisterRequiredBinaryExtractors(IWindsorContainer container, Assembly[] assembliesToRegister)
        {
            foreach (var assembly in assembliesToRegister)
            {
                container.Register(Classes.FromAssembly(assembly)
                    .IncludeNonPublicTypes()
                    .BasedOn<IRequiredBinaryFilesFromFileNameExtractor>()
                    .WithServiceBase()
                    .LifestyleSingleton());
            }
        }

        private void RegisterRefactorings(IWindsorContainer container, Assembly[] assembliesToRegister)
        {
            foreach (var assembly in assembliesToRegister)
            {
                container.Register(Classes.FromAssembly(assembly)
                    .IncludeNonPublicTypes()
                    .BasedOn<IRefactoring>()
                    .WithServiceSelf()
                    .LifestyleSingleton());
            }
        }

        private void RegisterCodeMetrics(IWindsorContainer container, Assembly[] assembliesToRegister)
        {
            foreach (var assembly in assembliesToRegister)
            {
                container.Register(Types.FromAssembly(assembly)
                    .IncludeNonPublicTypes()
                    .BasedOn<CodeMetric>()
                    .Unless(t => t == typeof(CodeMetric))
                    .WithServiceBase()
                    .LifestyleSingleton());
            }
            container.Register(Component.For<ICodeMetricsAnalyst>()
                .ImplementedBy<CodeMetricsAnalyst>()
                .LifestyleSingleton());
        }

        private void RegisterUnitTestingComSide(IWindsorContainer container)
        {
            container.Register(Component.For<IFakesFactory>()
                .ImplementedBy<FakesProviderFactory>()
                .LifestyleSingleton());
        }

        // note: settings namespace classes are injected in singleton scope
        private void RegisterConfiguration(IWindsorContainer container, Assembly[] assembliesToRegister)
        {
            var experimentalTypes = new List<Type>();
            foreach (var assembly in assembliesToRegister)
            {
                container.Register(Classes.FromAssembly(assembly)
                    .IncludeNonPublicTypes()
                    .BasedOn(typeof(ConfigurationServiceBase<>))
                    .WithServiceSelect((type, hierarchy) =>
                    {
                        // select closed generic interface
                        return type.GetInterfaces().Where(iface => iface.IsGenericType 
                            && iface.GetGenericTypeDefinition() == typeof(IConfigurationService<>));
                    })
                    .LifestyleSingleton());

                experimentalTypes.AddRange(assembly.GetTypes()
                    .Where(t => Attribute.IsDefined(t, typeof(ExperimentalAttribute))));
            }

            container.Register(Component.For(typeof(IDefaultSettings<>))
                .ImplementedBy(typeof(DefaultSettings<,>), new FixedGenericAppender(new[] { typeof(Properties.Settings) }))
                .IsFallback()
                .LifestyleTransient());

            var provider = new ExperimentalTypesProvider(experimentalTypes);
            container.Register(Component.For(typeof(IExperimentalTypesProvider))
                .DependsOn(Dependency.OnComponent<ViewModelBase, GeneralSettingsViewModel>())
                .LifestyleSingleton()
                .Instance(provider));

            container.Register(Component.For<IComProjectSerializationProvider>()
                .ImplementedBy<XmlComProjectSerializer>()
                .LifestyleTransient());
            container.Register(Component.For(typeof(IPersistenceService<>))
                .ImplementedBy(typeof(XmlPersistenceService<>))
                .LifestyleSingleton());

            container.Register(Component.For(typeof(IPersistenceService<ReferenceSettings>))
                .ImplementedBy(typeof(XmlContractPersistenceService<>))
                .LifestyleSingleton());

            container.Register(Component.For(typeof(IConfigurationService<Configuration>))
                .ImplementedBy(typeof(ConfigurationLoader))
                .LifestyleSingleton());
        }

        private void ApplyDefaultInterfaceConvention(IWindsorContainer container, Assembly[] assembliesToRegister)
        {
            foreach (var assembly in assembliesToRegister)
            {
                container.Register(Classes.FromAssembly(assembly)
                    .IncludeNonPublicTypes()
                    .Where(type => type.Namespace != null
                            && !type.Namespace.StartsWith("Rubberduck.VBEditor.SafeComWrappers")
                            && !type.Name.Equals(nameof(SelectionChangeService))
                            && !type.Name.Equals(nameof(AutoCompleteService))
                            && !type.Name.EndsWith("Factory")
                            && !type.Name.EndsWith("ConfigProvider")
                            && !type.Name.EndsWith("FakesProvider")
                            && !type.GetInterfaces().Contains(typeof(IInspection))
                            && type.NotDisabledOrExperimental(_initialSettings))
                    .WithService.DefaultInterfaces()
                    .LifestyleTransient()
                );
            }
        }

        private void RegisterFactories(IWindsorContainer container, Assembly[] assembliesToRegister)
        {
            foreach (var assembly in assembliesToRegister)
            {
                container.Register(Types.FromAssembly(assembly)
                    .IncludeNonPublicTypes()
                    .Where(type => type.IsInterface 
                                   && type.Name.EndsWith("Factory") 
                                   && !type.Name.Equals("IFakesFactory")
                                   && !type.Name.Equals("IAnnotationFactory")
                                   && type.NotDisabledOrExperimental(_initialSettings))
                    .WithService.Self()
                    .Configure(c => c.AsFactory())
                    .LifestyleSingleton());
            }
        }

        private void RegisterSourceCodeHandlers(IWindsorContainer container)
        {
            container.Register(Component.For<ISourceCodeHandler>()
                .ImplementedBy<ComponentSourceCodeHandlerSourceCodeHandlerAdapter>()
                .DependsOn(Dependency.OnComponent<IComponentSourceCodeHandler, CodeModuleComponentSourceCodeHandler>())
                .LifestyleSingleton()
                .Named("CodeModuleSourceCodeHandler"));
            container.Register(Component.For<ISourceCodeHandler>()
                .ImplementedBy<ComponentSourceCodeHandlerSourceCodeHandlerAdapter>()
                .DependsOn(Dependency.OnComponent<IComponentSourceCodeHandler, SourceFileHandlerComponentSourceCodeHandlerAdapter>())
                .LifestyleSingleton()
                .Named("SourceFileSourceCodeHandler"));
        }

        private void RegisterSpecialFactories(IWindsorContainer container)
        {
            container.Register(Component.For<ICodePaneHandler>()
                .ImplementedBy<CodePaneHandler>()
                .LifestyleSingleton());
            container.Register(Component.For<IFileSystemBrowserFactory>()
                .ImplementedBy<DialogFactory>()
                .LifestyleSingleton());
            container.Register(Component.For<IModuleRewriterFactory>()
                .ImplementedBy<ModuleRewriterFactory>()
                .DependsOn(Dependency.OnComponent("codePaneSourceCodeHandler", "CodeModuleSourceCodeHandler"),
                    Dependency.OnComponent("attributesSourceCodeHandler", "SourceFileSourceCodeHandler"))
                .LifestyleSingleton());
            container.Register(Component.For<IRubberduckParserErrorListenerFactory>()
                .ImplementedBy<ExceptionErrorListenerFactory>()
                .LifestyleSingleton());
            container.Register(Component.For<IParsePassErrorListenerFactory>()
                .ImplementedBy<MainParseErrorListenerFactory>()
                .LifestyleSingleton());
            container.Register(Component.For<PreprocessingParseErrorListenerFactory>()
                .ImplementedBy<PreprocessingParseErrorListenerFactory>()
                .LifestyleSingleton());
            container.Register(Component.For<IRewriteSessionFactory>()
                .ImplementedBy<RewriteSessionFactory>()
                .LifestyleSingleton());
            container.Register(Component.For<IAddRemoveReferencesPresenterFactory>()
                .ImplementedBy<AddRemoveReferencesPresenterFactory>()
                .LifestyleSingleton());
            container.Register(Component.For<IAnnotationArgumentViewModelFactory>()
                .ImplementedBy<AnnotationArgumentViewModelFactory>()
                .LifestyleSingleton());
            RegisterUnreachableCaseFactories(container);
        }

        private void RegisterUnreachableCaseFactories(IWindsorContainer container)
        {
            container.Register(Component.For<IParseTreeValueFactory>()
                .ImplementedBy<ParseTreeValueFactory>()
                .LifestyleSingleton());
        }


        private void RegisterQuickFixes(IWindsorContainer container, Assembly[] assembliesToRegister)
        {
            foreach (var assembly in assembliesToRegister)
            {
                container.Register(Classes.FromAssembly(assembly)
                    .IncludeNonPublicTypes()
                    .BasedOn<IQuickFix>()
                    .If(type => type.NotDisabledOrExperimental(_initialSettings))
                    .WithService.Base() 
                    .LifestyleSingleton());
            }
        }

        private void RegisterInspections(IWindsorContainer container, Assembly[] assembliesToRegister)
        {
            foreach (var assembly in assembliesToRegister)
            {
                container.Register(Classes.FromAssembly(assembly)
                    .IncludeNonPublicTypes()
                    .BasedOn<IInspection>()
                    .If(type => type.NotDisabledOrExperimental(_initialSettings))
                    .WithService.Base()
                    .LifestyleTransient());
            }
        }

        private void RegisterAutoCompletes(IWindsorContainer container, Assembly[] assembliesToRegister)
        {
            foreach (var assembly in assembliesToRegister)
            {
                container.Register(Classes.FromAssembly(assembly)
                    .IncludeNonPublicTypes()
                    .BasedOn<AutoCompleteHandlerBase>()
                    .If(type => type.NotDisabledOrExperimental(_initialSettings))
                    .WithService.Base()
                    .LifestyleTransient());
            }
        }

        private void RegisterParseTreeInspections(IWindsorContainer container, Assembly[] assembliesToRegister)
        {
            foreach (var assembly in assembliesToRegister)
            {
                container.Register(Classes.FromAssembly(assembly)
                    .IncludeNonPublicTypes()
                    .BasedOn<IParseTreeInspection>()
                    .If(type => type.NotDisabledOrExperimental(_initialSettings))
                    .WithService.Select(new[] { typeof(IInspection) })
                    .LifestyleTransient());
            }
        }

        private void RegisterRubberduckMenu(IWindsorContainer container)
        {
            if (!_addin.CommandBarLocations.TryGetValue(CommandBarSite.MenuBar, out var location))
            {
                return;
            }

            var controls = MainCommandBarControls(location.ParentId);
            var beforeIndex = FindRubberduckMenuInsertionIndex(controls, location.BeforeControlId);
            var menuItemTypes = RubberduckMenuItems();
            RegisterMenu<RubberduckParentMenu>(container, controls, beforeIndex, menuItemTypes);
        }

        private void RegisterMenu<TMenu>(IWindsorContainer container, ICommandBarControls controls, int beforeIndex, Type[] menuItemTypes) where TMenu : IParentMenuItem
        {
            var nonExperimentalMenuItems = menuItemTypes.Where(type => type.NotDisabledOrExperimental(_initialSettings)).ToArray();
            container.Register(Component.For<IParentMenuItem>()
                .ImplementedBy<TMenu>()
                .LifestyleTransient()
                .DependsOn(
                    Dependency.OnValue<int>(beforeIndex),
                    Dependency.OnComponentCollection<IEnumerable<IMenuItem>>(nonExperimentalMenuItems))
                .OnCreate((kernel, item) => item.Parent = controls));
        }

        private Type[] RubberduckMenuItems()
        {
            return new[]
            {
                typeof(RefreshCommandMenuItem),
                typeof(AboutCommandMenuItem),
                typeof(SettingsCommandMenuItem),
                typeof(InspectionResultsCommandMenuItem),
                typeof(UnitTestingParentMenu),
                typeof(SmartIndenterParentMenu),
                typeof(ToolsParentMenu),
                typeof(RefactoringsParentMenu),
                typeof(NavigateParentMenu)
            };
        }

        private int FindRubberduckMenuInsertionIndex(ICommandBarControls controls, int beforeId)
        {
            for (var i = 1; i <= controls.Count; i++)
            {
                using (var item = controls[i])
                {
                    if (item.IsBuiltIn && item.Id == beforeId)
                    {
                        return i;
                    }
                }
            }

            return controls.Count;
        }

        private ICommandBarControls MainCommandBarControls(int commandBarIndex)
        {
            ICommandBarControls controls;
            using (var commandBars = _vbe.CommandBars)
            {
                using (var menuBar = commandBars[commandBarIndex])
                {
                    controls = menuBar.Controls;
                }
            }
            return controls;
        }

        private void RegisterCodePaneContextMenu(IWindsorContainer container)
        {
            if (!_addin.CommandBarLocations.TryGetValue(CommandBarSite.CodeWindow, out var location))
            {
                return;
            }

            var controls = MainCommandBarControls(location.ParentId);
            var beforeIndex = FindRubberduckMenuInsertionIndex(controls, location.BeforeControlId);
            var menuItemTypes = CodePaneContextMenuItems();
            RegisterMenu<CodePaneContextParentMenu>(container, controls, beforeIndex, menuItemTypes);
        }

        private Type[] CodePaneContextMenuItems()
        {
            return new Type[]
            {
                typeof(CodePaneRefactoringsParentMenu),
                typeof(AnnotateParentMenu),
                typeof(SmartIndenterParentMenu),
                typeof(FindSymbolCommandMenuItem),
                typeof(FindAllReferencesCommandMenuItem),
                typeof(FindAllImplementationsCommandMenuItem),
                typeof(RunSelectedTestModuleCommandMenuItem),
                typeof(RunSelectedTestMethodCommandMenuItem)
            };
        }

        private void RegisterFormDesignerContextMenu(IWindsorContainer container)
        {
            if (!_addin.CommandBarLocations.TryGetValue(CommandBarSite.MsForm, out var location))
            {
                return;
            }

            var controls = MainCommandBarControls(location.ParentId);
            var beforeIndex = FindRubberduckMenuInsertionIndex(controls, location.BeforeControlId);
            var menuItemTypes = FormDesignerContextMenuItems();
            RegisterMenu<FormDesignerContextParentMenu>(container, controls, beforeIndex, menuItemTypes);
        }

        private Type[] FormDesignerContextMenuItems()
        {
            return new Type[]
            {
                typeof(FormDesignerRefactorRenameCommandMenuItem),
                typeof(FormDesignerFindAllReferencesCommandMenuItem)
            };
        }

        private void RegisterFormDesignerControlContextMenu(IWindsorContainer container)
        {
            if (!_addin.CommandBarLocations.TryGetValue(CommandBarSite.MsFormControl, out var location))
            {
                return;
            }

            var controls = MainCommandBarControls(location.ParentId);
            var beforeIndex = FindRubberduckMenuInsertionIndex(controls, location.BeforeControlId);
            var menuItemTypes = FormDesignerContextMenuItems();
            RegisterMenu<FormDesignerControlContextParentMenu>(container, controls, beforeIndex, menuItemTypes);
        }

        private void RegisterProjectExplorerContextMenu(IWindsorContainer container)
        {
            if (!_addin.CommandBarLocations.TryGetValue(CommandBarSite.ProjectExplorer, out var location))
            {
                return;
            }

            var controls = MainCommandBarControls(location.ParentId);
            var beforeIndex = FindRubberduckMenuInsertionIndex(controls, location.BeforeControlId);
            var menuItemTypes = ProjectWindowContextMenuItems();
            RegisterMenu<ProjectWindowContextParentMenu>(container, controls, beforeIndex, menuItemTypes);
        }

        private Type[] ProjectWindowContextMenuItems()
        {
            return new[]
            {
                typeof(ProjectExplorerRefactorRenameCommandMenuItem),
                typeof(FindSymbolCommandMenuItem),
                typeof(FindAllReferencesCommandMenuItem),
                typeof(FindAllImplementationsCommandMenuItem),
                typeof(ProjectExplorerAddRemoveReferencesCommandMenuItem)
            };
        }

        private void RegisterRubberduckCommandBar(IWindsorContainer container)
        {
            container.Register(Component.For<RubberduckCommandBar>()
                .LifestyleTransient()
                .DependsOn(Dependency.OnComponentCollection<IEnumerable<ICommandMenuItem>>(RubberduckCommandBarItems()))
                .OnCreate((kernel, item) => item.Parent = kernel.Resolve<ICommandBars>()));
        }

        private Type[] RubberduckCommandBarItems()
        {
            var types = new List<Type>
            {
                typeof(ReparseCommandMenuItem),
                typeof(ShowParserErrorsCommandMenuItem),
                typeof(ContextSelectionLabelMenuItem),
                typeof(ContextDescriptionLabelMenuItem),
                typeof(ReferenceCounterLabelMenuItem)
            };

            AttachRubberduckDebugCommandBarItems(ref types);

            return types.ToArray();
        }

        [Conditional("DEBUG")]
        private static void AttachRubberduckDebugCommandBarItems(ref List<Type> types)
        {
            types.Add(typeof(SerializeProjectsCommandMenuItem));
        }

        private void RegisterParentMenus(IWindsorContainer container)
        {
            RegisterParentMenu<UnitTestingParentMenu>(container, UnitTestingMenuItems());
            RegisterParentMenu<RefactoringsParentMenu>(container, RefactoringsMenuItems());
            RegisterParentMenu<CodePaneRefactoringsParentMenu>(container, RefactoringsMenuItems());
            RegisterParentMenu<NavigateParentMenu>(container, NavigateMenuItems());
            RegisterParentMenu<SmartIndenterParentMenu>(container, SmartIndenterMenuItems());
            RegisterParentMenu<AnnotateParentMenu>(container, AnnotateMenuItems());
            RegisterParentMenu<ToolsParentMenu>(container, ToolsMenuItems());
        }

        private void RegisterParentMenu<TParentMenu>(IWindsorContainer container, Type[] menuItemTypes) where TParentMenu : IParentMenuItem
        {
            var nonExperimentalMenuItems = menuItemTypes.Where(type => type.NotDisabledOrExperimental(_initialSettings)).ToArray();
            container.Register(Component.For<IMenuItem, TParentMenu>()
                .ImplementedBy<TParentMenu>()
                .LifestyleTransient()
                .DependsOn(Dependency.OnComponentCollection<IEnumerable<IMenuItem>>(nonExperimentalMenuItems)));
        }

        private Type[] UnitTestingMenuItems()
        {
            return new[]
            {
                typeof(RunAllTestsCommandMenuItem),
                typeof(TestExplorerCommandMenuItem),
                typeof(AddTestModuleCommandMenuItem),
                typeof(AddTestMethodCommandMenuItem),
                typeof(AddTestMethodExpectedErrorCommandMenuItem)
            };
        }

        private Type[] RefactoringsMenuItems()
        {
            return new[]
            {
                typeof(CodePaneRefactorRenameCommandMenuItem),
                typeof(RefactorExtractMethodCommandMenuItem),
                typeof(RefactorReorderParametersCommandMenuItem),
                typeof(RefactorRemoveParametersCommandMenuItem),
                typeof(RefactorIntroduceParameterCommandMenuItem),
                typeof(RefactorIntroduceFieldCommandMenuItem),
                typeof(RefactorEncapsulateFieldCommandMenuItem),
                typeof(RefactorMoveCloserToUsageCommandMenuItem),
                typeof(RefactorExtractInterfaceCommandMenuItem),
                typeof(RefactorImplementInterfaceCommandMenuItem),
                typeof(CodePaneRefactorMoveToFolderCommandMenuItem),
                typeof(CodePaneRefactorMoveContainingFolderCommandMenuItem)
            };
        }

        private Type[] NavigateMenuItems()
        {
            return new[]
            {
                typeof(CodeExplorerCommandMenuItem),
                typeof(RegexSearchReplaceCommandMenuItem),               
                typeof(FindSymbolCommandMenuItem),
                typeof(FindAllReferencesCommandMenuItem),
                typeof(FindAllImplementationsCommandMenuItem)
            };
        }

        private Type[] SmartIndenterMenuItems()
        {
            return new[]
            {
                typeof(IndentCurrentProcedureCommandMenuItem),
                typeof(IndentCurrentModuleCommandMenuItem),
                typeof(IndentCurrentProjectCommandMenuItem),
                typeof(NoIndentAnnotationCommandMenuItem)
            };
        }

        private Type[] AnnotateMenuItems()
        {
            return new[]
            {
                typeof(AnnotateSelectedDeclarationCommandMenuItem),
                typeof(AnnotateSelectedModuleCommandMenuItem),
                typeof(AnnotateSelectedMemberCommandMenuItem)
            };
        }

        private Type[] ToolsMenuItems()
        {
            var items = new List<Type>
            {
                typeof(RegexAssistantCommandMenuItem),
                typeof(ToDoExplorerCommandMenuItem),
                typeof(CodeMetricsCommandMenuItem),
                typeof(ExportAllCommandMenuItem),
                typeof(ToolMenuAddRemoveReferencesCommandMenuItem)
            };
            
            return items.ToArray();
        }

        private void RegisterSettingsViewModel(IWindsorContainer container)
        {
            container.Register(Types
                .FromAssemblyInThisApplication()
                .IncludeNonPublicTypes()
                .BasedOn(typeof(SettingsViewModelBase<>))
                .LifestyleTransient()
                .WithServiceSelect((type, types) =>
                {
                    var face = type.GetInterfaces().FirstOrDefault(i =>
                        i.IsGenericType && i.GetGenericTypeDefinition() == typeof(ISettingsViewModel<>));

                    return face == null ? new[] { type } : new[] { type, face };
                })
            );
        }

        private void RegisterRefactoringPreviewProviders(IWindsorContainer container)
        {
            container.Register(Types
                .FromAssemblyInThisApplication()
                .IncludeNonPublicTypes()
                .BasedOn(typeof(IRefactoringPreviewProvider<>))
                .LifestyleSingleton()
                .WithServiceSelect((type, types) =>
                {
                    var face = type.GetInterfaces().FirstOrDefault(i =>
                        i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IRefactoringPreviewProvider<>));

                    return face == null ? new[] { type } : new[] { type, face };
                })
            );
        }

        private void RegisterRefactoringDialogs(IWindsorContainer container)
        {
            container.Register(Types
                .FromAssemblyInThisApplication()
                .IncludeNonPublicTypes()
                .BasedOn(typeof(IRefactoringView<>))
                .LifestyleTransient()
                .WithServiceSelect((type, types) =>
                {
                    var face = type.GetInterfaces().FirstOrDefault(i =>
                        i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IRefactoringView<>));

                    return face == null ? new[] { type } : new[] { type, face };
                })
            );
            container.Register(Types
                .FromAssemblyInThisApplication()
                .IncludeNonPublicTypes()
                .BasedOn(typeof(IRefactoringViewModel<>))
                .LifestyleTransient()
                .WithServiceSelect((type, types) =>
                {
                    var face = type.GetInterfaces().FirstOrDefault(i =>
                        i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IRefactoringViewModel<>));

                    return face == null ? new[] { type } : new[] {type, face};
                })
            );
            container.Register(Types
                .FromAssemblyInThisApplication()
                .IncludeNonPublicTypes()
                .BasedOn(typeof(IRefactoringDialog<,,>))
                .LifestyleTransient()
                .WithServiceSelect((type, types) =>
                {
                    var face = type.GetInterfaces().FirstOrDefault(i =>
                        i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IRefactoringDialog<,,>));

                    if (face == null)
                    {
                        return new[] { type };
                    }

                    var model = face.GenericTypeArguments[0];

                    var view = face.GenericTypeArguments[1];
                    var interfaceView = view.GetInterfaces().FirstOrDefault(i =>
                        i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IRefactoringView<>));

                    if (interfaceView == null)
                    {
                        return new[] { type };
                    }
                    
                    var viewModel = face.GenericTypeArguments[2];
                    var interfaceViewModel = viewModel.GetInterfaces().FirstOrDefault(i =>
                        i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IRefactoringViewModel<>));

                    if (interfaceViewModel == null)
                    {
                        return new[] {type};
                    }

                    var closedFace = typeof(IRefactoringDialog<,,>).MakeGenericType(model, interfaceView, interfaceViewModel);

                    return new[] { type, closedFace };
                })
            );
        }

        private void RegisterCommandMenuItems(IWindsorContainer container)
        {
            //note: The name of a registration is the full name of the implementation if not specified otherwise.
            container.Register(Classes.FromAssemblyContaining<ICommandMenuItem>()
                .IncludeNonPublicTypes()
                .BasedOn<ICommandMenuItem>()
                .If(type => type.NotDisabledOrExperimental(_initialSettings))
                .WithService.Base()
                .LifestyleTransient());
        }

        private void RegisterCommands(IWindsorContainer container)
        {
            // assumption: All Commands are in the same assembly as CommandBase
            container.Register(Classes.FromAssemblyContaining(typeof(CommandBase))
                .IncludeNonPublicTypes()
                .Where(type => type.IsBasedOn(typeof(CommandBase))
                    && type != typeof(DelegateCommand) // DelegateCommand is not intended to be injected!
                    && type.NotDisabledOrExperimental(_initialSettings))
                .WithService.Select(new[] { typeof(CommandBase) })
                .WithService.Self()
                .WithService.DefaultInterfaces()
                .LifestyleTransient());
        }

        private void RegisterSmartIndenter(IWindsorContainer container)
        {
            container.Register(Component.For<IIndenter, Indenter>()
                .ImplementedBy<Indenter>()
                .LifestyleSingleton());
            container.Register(Component.For<IIndenterSettings>()
                .ImplementedBy<IndenterSettings>()
                .LifestyleSingleton());
            container.Register(Component.For<Func<IIndenterSettings>>()
                .UsingFactoryMethod(kernel => (Func<IIndenterSettings>)(() => kernel.Resolve<IConfigurationService<Configuration>>()
                   .Read().UserSettings.IndenterSettings))
                .LifestyleTransient()); //todo: clean up this registration
        }

        private void RegisterWindowsHooks(IWindsorContainer container)
        {
            using (var mainWindow = _vbe.MainWindow)
            {
                var mainWindowHwnd = (IntPtr)mainWindow.HWnd;
                container.Register(Component.For<IRubberduckHooks>()
                    .ImplementedBy<RubberduckHooks>()
                    .DependsOn(Dependency.OnValue<IntPtr>(mainWindowHwnd))
                    .LifestyleSingleton());
            }
        }
        
        private void RegisterDockableUserControls(IWindsorContainer container)
        {
            container.Register(Classes.FromAssemblyContaining<IDockableUserControl>()
                .IncludeNonPublicTypes()
                .BasedOn<IDockableUserControl>()
                .LifestyleSingleton());
        }

        private void RegisterDockablePresenters(IWindsorContainer container)
        {
            container.Register(Classes.FromAssemblyContaining<IDockablePresenter>()
                .IncludeNonPublicTypes()
                .BasedOn<IDockablePresenter>()
                .WithServiceSelf()
                .WithServices(new[] { typeof(IDockablePresenter) })
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

        private void OverridePropertyInjection(IWindsorContainer container)
        {
            // remove default property injection
            var propInjector = container.Kernel.ComponentModelBuilder
                .Contributors
                .OfType<PropertiesDependenciesModelInspector>()
                .Single();
            container.Kernel.ComponentModelBuilder.RemoveContributor(propInjector);

            container.Kernel.ComponentModelBuilder.AddContributor(new RubberduckPropertiesInspector());
        }

        private void RegisterParsingEngine(IWindsorContainer container)
        {
            RegisterCustomDeclarationLoadersToParser(container);
            RegisterAnnotationProcessing(container);

            container.Register(Component.For<ICompilationArgumentsProvider, ICompilationArgumentsCache>()
                .ImplementedBy<CompilationArgumentsCache>()
                .DependsOn(Dependency.OnComponent<ICompilationArgumentsProvider,CompilationArgumentsProvider>())
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
            container.Register(Component.For<IComLibraryProvider>()
                .ImplementedBy<ComLibraryProvider>()
                .LifestyleSingleton());
            container.Register(Component.For<IReferencedDeclarationsCollector>()
                .ImplementedBy<LibraryReferencedDeclarationsCollector>()
                .LifestyleSingleton());
            container.Register(Component.For<ITokenStreamPreprocessor>()
                .ImplementedBy<VBAPreprocessor>()
                .DependsOn(Dependency.OnComponent<ITokenStreamParser, VBAPreprocessorParser>())
                .LifestyleSingleton());
            container.Register(Component.For<VBAPredefinedCompilationConstants>()
                .ImplementedBy<VBAPredefinedCompilationConstants>()
                .DependsOn(Dependency.OnValue<double>(double.Parse(_vbe.Version, CultureInfo.InvariantCulture)))
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
            container.Register(Component.For<ITokenStreamParser>()
                .ImplementedBy<VBATokenStreamParser>()
                .LifestyleSingleton());
            container.Register(Component.For<IModuleParser>()
                .ImplementedBy<ModuleParser>()
                .DependsOn(Dependency.OnComponent("codePaneSourceCodeProvider", "CodeModuleSourceCodeHandler"),
                    Dependency.OnComponent("attributesSourceCodeProvider", "SourceFileSourceCodeHandler"))
                .LifestyleSingleton());
            container.Register(Component.For<ITypeLibWrapperProvider>()
                .ImplementedBy<TypeLibWrapperProvider>()
                .LifestyleSingleton());
            container.Register(Component.For<IUserComProjectRepository, IUserComProjectProvider>()
                .ImplementedBy<UserProjectRepository>()
                .LifestyleSingleton());
            container.Register(Component.For<IDeclarationsFromComProjectLoader>()
                .ImplementedBy<DeclarationsFromComProjectLoader>()
                .LifestyleSingleton());
            container.Register(Component.For<IUserComProjectSynchronizer>()
                .ImplementedBy<UserComProjectSynchronizer>()
                .LifestyleSingleton());
            container.Register(Component.For<IProjectsToResolveFromComProjectSelector>()
                .ImplementedBy<ProjectsToResolveFromComProjectsSelector>()
                .LifestyleSingleton());
        }

        private void RegisterAnnotationProcessing(IWindsorContainer container)
        {
            foreach (Assembly referenced in AssembliesToRegister())
            {
                container.Register(Classes.FromAssembly(referenced)
                    .IncludeNonPublicTypes()
                    .BasedOn<IAnnotation>()
                    .WithServiceAllInterfaces()
                    .LifestyleSingleton());
            }
            container.Register(Component.For<IAnnotationFactory>()
                .ImplementedBy<VBAParserAnnotationFactory>()
                .LifestyleSingleton());
        }

        private void RegisterTypeLibApi(IWindsorContainer container)
        {
            container.Register(Component.For<IVBETypeLibsAPI>()
                .ImplementedBy<VBETypeLibsAPI>()
                .LifestyleSingleton());
        }

        private void RegisterCustomDeclarationLoadersToParser(IWindsorContainer container)
        {
            container.Register(Classes.FromAssemblyContaining<ICustomDeclarationLoader>()
                .BasedOn<ICustomDeclarationLoader>()
                .WithService.Base()
                .LifestyleSingleton());
        }

        //note: We assume that the full names of all assemblies belonging to Rubberduck start with 'Rubberduck'.
        public IEnumerable<Assembly> AssembliesToRegister()
        {
            return GetDistinctTransitivelyReferencedAssemblies(Assembly.GetExecutingAssembly(), name => name.FullName.StartsWith("Rubberduck"))
                //For some reason the inspections assembly is not referenced transitively.
                .Concat(new [] { Assembly.GetAssembly(typeof(Inspector)) })
                //Theoretically we shouldn't have anything to register here, but better safe than sorry.
                .Concat(new[] { Assembly.GetExecutingAssembly() })
                .Distinct();
        }

        /// <summary>
        /// Recursively finds all assemblies referenced by the <parameref name="assembly"/> directly or indirectly through a path of assemblies satisfying the <paramref name="filterPredicate"/>.
        /// </summary>
        /// <param name="assembly">The assembly for which to find all transitive references</param>
        /// <param name="filterPredicate">Filter to restrict the assemblies considered</param>
        private IEnumerable<Assembly> GetDistinctTransitivelyReferencedAssemblies(Assembly assembly, Predicate<AssemblyName> filterPredicate)
        {
            var referencedAssemblies = assembly.GetReferencedAssemblies()
                                                .Where(asmbly => filterPredicate(asmbly))
                                                .Distinct().Select(Assembly.Load)
                                                .ToList();
            //This terminates because circular assembly references are illegal.
            var transitiveReferences = referencedAssemblies.SelectMany(asmbly => GetDistinctTransitivelyReferencedAssemblies(asmbly, filterPredicate)).ToList();
            referencedAssemblies.AddRange(transitiveReferences);
            return referencedAssemblies.Distinct();
        }

        private void RegisterAppWithSpecialDependencies(IWindsorContainer container)
        {
            container.Register(Component.For<CommandBase>()
                .ImplementedBy<VersionCheckCommand>()
                .Named(nameof(VersionCheckCommand))
                .LifestyleSingleton());
            container.Register(Component.For<App>()
                .DependsOn(Dependency.OnComponent<CommandBase, VersionCheckCommand>())
                .LifestyleSingleton());
        }

        private void RegisterInstances(IWindsorContainer container)
        {
            container.Register(Component.For<IVBE>().Instance(_vbe));
            container.Register(Component.For<IAddIn>().Instance(_addin));
            //note: This registration makes Castle Windsor inject _vbe_CommandBars in all ICommandBars Parent properties.
            container.Register(Component.For<ICommandBars>().Instance(_vbe.CommandBars));
            container.Register(Component.For<IUiContextProvider>().Instance(UiContextProvider.Instance()).LifestyleSingleton());
            container.Register(Component.For<IVbeEvents>().Instance(VbeEvents.Initialize(_vbe)).LifestyleSingleton());
            container.Register(Component.For<ITempSourceFileHandler>().Instance(_vbe.TempSourceFileHandler).LifestyleSingleton());
            container.Register(Component.For<IPersistencePathProvider>().Instance(PersistencePathProvider.Instance).LifestyleSingleton());
            container.Register(Component.For<IVbeNativeApi>().Instance(_vbeNativeApi).LifestyleSingleton());
            container.Register(Component.For<IBeepInterceptor>().Instance(_beepInterceptor).LifestyleSingleton());
        }
    }
}
