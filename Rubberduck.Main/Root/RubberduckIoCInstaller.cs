﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Castle.Facilities.TypedFactory;
using Castle.MicroKernel.ModelBuilder.Inspectors;
using Castle.MicroKernel.Registration;
using Castle.MicroKernel.Resolvers.SpecializedResolvers;
using Castle.MicroKernel.SubSystems.Configuration;
using Castle.Windsor;
using Rubberduck.ComClientLibrary.UnitTesting;
using Rubberduck.Common;
using Rubberduck.Common.Hotkeys;
using Rubberduck.Inspections.Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.Symbols.DeclarationLoaders;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;
using Rubberduck.UI;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.Command.MenuItems.CommandBars;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Controls;
using Rubberduck.UI.Refactorings;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.UI.UnitTesting;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Component = Castle.MicroKernel.Registration.Component;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.Parsing.Common;
using Rubberduck.VBEditor.ComManagement.TypeLibsAPI;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.Utility;
using Rubberduck.AutoComplete;
using Rubberduck.AutoComplete.Service;
using Rubberduck.CodeAnalysis.CodeMetrics;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA.ComReferenceLoading;
using Rubberduck.Parsing.VBA.DeclarationResolving;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Parsing.VBA.ReferenceManagement;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using Rubberduck.VBEditor.SourceCodeHandling;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA.Parsing.ParsingExceptions;

namespace Rubberduck.Root
{
    public class RubberduckIoCInstaller : IWindsorInstaller
    {
        private readonly IVBE _vbe;
        private readonly IAddIn _addin;
        private readonly GeneralSettings _initialSettings;

        public RubberduckIoCInstaller(IVBE vbe, IAddIn addin, GeneralSettings initialSettings)
        {
            _vbe = vbe;
            _addin = addin;
            _initialSettings = initialSettings;
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

            RegisterConstantVbeAndAddIn(container);
            RegisterAppWithSpecialDependencies(container);
            RegisterUnitTestingComSide(container);

            container.Register(Component.For<Version>()
                     .UsingFactoryMethod(() => Assembly.GetExecutingAssembly().GetName().Version)
                     .LifestyleSingleton());

            container.Register(Component.For<IProjectsProvider, IProjectsRepository>()
                .ImplementedBy<ProjectsRepository>()
                .LifestyleSingleton());
            container.Register(Component.For<RubberduckParserState, IParseTreeProvider, IDeclarationFinderProvider>()
                .ImplementedBy<RubberduckParserState>()
                .LifestyleSingleton());
            container.Register(Component.For<ISelectionChangeService>()
                .ImplementedBy<SelectionChangeService>()
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
            RegisterParsingEngine(container);
            RegisterTypeLibApi(container);

            container.Register(Component.For<TestExplorerModel>()
                .LifestyleSingleton());
            container.Register(Component.For<IVBEInteraction>()
                .ImplementedBy<VBEInteraction>()
                .LifestyleSingleton());

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

            RegisterParseTreeInspections(container, assembliesToRegister);
            RegisterInspections(container, assembliesToRegister);
            RegisterQuickFixes(container, assembliesToRegister);
            RegisterAutoCompletes(container, assembliesToRegister);
            RegisterCodeMetrics(container, assembliesToRegister);

            RegisterSpecialFactories(container);
            RegisterFactories(container, assembliesToRegister);

            ApplyDefaultInterfaceConvention(container, assembliesToRegister);
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
                    .Where(type => type.Namespace == typeof(Configuration).Namespace && type.NotDisabledOrExperimental(_initialSettings))
                    .WithService.AllInterfaces()
                    .LifestyleSingleton());

                experimentalTypes.AddRange(assembly.GetTypes()
                    .Where(t => Attribute.IsDefined(t, typeof(ExperimentalAttribute))));
            }

            // FIXME correctly register experimentalFeatureTypes.
            // This is probably blocked until GeneralSettingsViewModel is no more newed up in SettingsForm's code-behind
            //container.Register(Component.For(typeof(IEnumerable<Type>))
            //    .DependsOn(Dependency.OnComponent<ViewModelBase, GeneralSettingsViewModel>())
            //    .LifestyleSingleton()
            //    .Instance(experimentalTypes));

            container.Register(Component.For<IComProjectSerializationProvider>()
                .ImplementedBy<XmlComProjectSerializer>()
                .LifestyleTransient());
            container.Register(Component.For(typeof(IPersistanceService<>), typeof(IFilePersistanceService<>))
                .ImplementedBy(typeof(XmlPersistanceService<>))
                .LifestyleSingleton());

            container.Register(Component.For<IConfigProvider<IndenterSettings>>()
                .ImplementedBy<IndenterConfigProvider>()
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
                                   && type.NotDisabledOrExperimental(_initialSettings))
                    .WithService.Self()
                    .Configure(c => c.AsFactory())
                    .LifestyleSingleton());
            }
        }

        private void RegisterSpecialFactories(IWindsorContainer container)
        {
            container.Register(Component.For<ICodePaneHandler>()
                .ImplementedBy<CodePaneSourceCodeHandler>()
                .LifestyleSingleton());
            container.Register(Component.For<IFolderBrowserFactory>()
                .ImplementedBy<DialogFactory>()
                .LifestyleSingleton());
            container.Register(Component.For<IModuleRewriterFactory>()
                .ImplementedBy<ModuleRewriterFactory>()
                .DependsOn(Dependency.OnComponent("codePaneSourceCodeHandler", typeof(CodePaneSourceCodeHandler)),
                    Dependency.OnComponent("attributesSourceCodeHandler", typeof(SourceFileHandlerSourceCodeHandlerAdapter)))
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
                typeof(RefactoringsParentMenu),
                typeof(SmartIndenterParentMenu),
                typeof(FindSymbolCommandMenuItem),
                typeof(FindAllReferencesCommandMenuItem),
                typeof(FindAllImplementationsCommandMenuItem)
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
                typeof(FindAllImplementationsCommandMenuItem)
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
            return new[]
            {
                typeof(ReparseCommandMenuItem),
                typeof(ShowParserErrorsCommandMenuItem),
                typeof(ContextSelectionLabelMenuItem),
                typeof(ContextDescriptionLabelMenuItem),
                typeof(ReferenceCounterLabelMenuItem),
#if DEBUG
                typeof(SerializeProjectsCommandMenuItem)
#endif
            };
        }

        private void RegisterParentMenus(IWindsorContainer container)
        {
            RegisterParentMenu<UnitTestingParentMenu>(container, UnitTestingMenuItems());
            RegisterParentMenu<RefactoringsParentMenu>(container, RefactoringsMenuItems());
            RegisterParentMenu<NavigateParentMenu>(container, NavigateMenuItems());
            RegisterParentMenu<SmartIndenterParentMenu>(container, SmartIndenterMenuItems());
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
                typeof(RefactorImplementInterfaceCommandMenuItem)
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

        private Type[] ToolsMenuItems()
        {
            var items = new List<Type>
            {
                typeof(RegexAssistantCommandMenuItem),
                typeof(ToDoExplorerCommandMenuItem),
                typeof(CodeMetricsCommandMenuItem),
                typeof(ExportAllCommandMenuItem)
            };
            
            return items.ToArray();
        }

        private void RegisterRefactoringDialogs(IWindsorContainer container)
        {
            container.Register(Component.For<IRefactoringDialog<RenameViewModel>>()
                .ImplementedBy<RenameDialog>()
                .LifestyleTransient());
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
                .UsingFactoryMethod(kernel => (Func<IIndenterSettings>)(() => kernel.Resolve<IGeneralConfigService>()
                   .LoadConfiguration().UserSettings.IndenterSettings))
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
                .DependsOn(Dependency.OnComponent("codePaneSourceCodeProvider", typeof(CodePaneSourceCodeHandler)),
                    Dependency.OnComponent("attributesSourceCodeProvider", typeof(SourceFileHandlerSourceCodeHandlerAdapter)))
                .LifestyleSingleton());
            container.Register(Component.For<ITypeLibWrapperProvider>()
                .ImplementedBy<TypeLibWrapperProvider>()
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

        private void RegisterConstantVbeAndAddIn(IWindsorContainer container)
        {
            container.Register(Component.For<IVBE>().Instance(_vbe));
            container.Register(Component.For<IAddIn>().Instance(_addin));
            //note: This registration makes Castle Windsor inject _vbe_CommandBars in all ICommandBars Parent properties.
            container.Register(Component.For<ICommandBars>().Instance(_vbe.CommandBars));
            container.Register(Component.For<IUiContextProvider>().Instance(UiContextProvider.Instance()).LifestyleSingleton());
            container.Register(Component.For<IVBEEvents>().Instance(VBEEvents.Initialize(_vbe)).LifestyleSingleton());
            container.Register(Component.For<ITempSourceFileHandler>().Instance(_vbe.TempSourceFileHandler));
        }
    }
}
