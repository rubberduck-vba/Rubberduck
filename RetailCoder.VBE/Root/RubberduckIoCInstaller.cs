using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using Castle.Facilities.TypedFactory;
using Castle.MicroKernel.ModelBuilder.Inspectors;
using Castle.MicroKernel.Registration;
using Castle.MicroKernel.Resolvers.SpecializedResolvers;
using Castle.MicroKernel.SubSystems.Configuration;
using Castle.Windsor;
using Rubberduck.Common;
using Rubberduck.Common.Hotkeys;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.Symbols.DeclarationLoaders;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;
using Rubberduck.SourceControl;
using Rubberduck.UI;
using Rubberduck.UI.CodeExplorer;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.Command.MenuItems.CommandBars;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;
using Rubberduck.UI.Controls;
using Rubberduck.UI.Inspections;
using Rubberduck.UI.Refactorings;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.UI.SourceControl;
using Rubberduck.UI.ToDoItems;
using Rubberduck.UI.UnitTesting;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;
using Component = Castle.MicroKernel.Registration.Component;
using Rubberduck.UI.CodeMetrics;
using Rubberduck.Parsing.Common;

namespace Rubberduck.Root
{
    public class RubberduckIoCInstaller : IWindsorInstaller
    {
        private readonly IVBE _vbe;
        private readonly IAddIn _addin;
        private readonly GeneralSettings _initialSettings;

        private const int MenuBar = 1;
        private const int CodeWindow = 9;
        private const int ProjectWindow = 14;
        private const int MsForms = 17;
        private const int MsFormsControl = 18;

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
            DeactivatePropertyInjection(container);

            RegisterConstantVbeAndAddIn(container);
            RegisterAppWithSpecialDependencies(container);

            container.Register(Component.For<RubberduckParserState, IParseTreeProvider, IDeclarationFinderProvider>()
                .ImplementedBy<RubberduckParserState>()
                .LifestyleSingleton());
            container.Register(Component.For<ISelectionChangeService>()
                .ImplementedBy<SelectionChangeService>()
                .LifestyleSingleton());
            container.Register(Component.For<ISourceControlProvider>()
                .ImplementedBy<GitProvider>()
                .LifestyleSingleton());
            //container.Register(Component.For<GitProvider>()
            //    .LifestyleSingleton());
            container.Register(Component.For<IOperatingSystem>()
                .ImplementedBy<WindowsOperatingSystem>()
                .LifestyleSingleton());

            container.Register(Component.For<DeclarationFinder>()
                .ImplementedBy<ConcurrentlyConstructedDeclarationFinder>()
                .LifestyleTransient());

            RegisterSmartIndenter(container);
            RegisterParsingEngine(container);

            container.Register(Component.For<TestExplorerModel>()
                .LifestyleSingleton());

            RegisterRefactoringDialogs(container);

            container.Register(Component.For<ISearchResultsWindowViewModel>()
                .ImplementedBy<SearchResultsWindowViewModel>()
                .LifestyleSingleton());
            container.Register(Component.For<SearchResultPresenterInstanceManager>()
                .LifestyleSingleton());

            RegisterSourceControlControls(container);

            RegisterDockablePresenters(container);
            RegisterDockableUserControls(container);

            RegisterCommands(container);
            RegisterCommandMenuItems(container);
            RegisterParentMenus(container);
            RegisterCodeExplorerViewModelWithCodeExplorerCommands(container);

            RegisterRubberduckCommandBar(container);
            RegisterRubberduckMenu(container);
            RegisterCodePaneContextMenu(container);
            RegisterFormDesignerContextMenu(container);
            RegisterFormDesignerControlContextMenu(container);
            RegisterProjectExplorerContextMenu(container);

            RegisterWindowsHooks(container);

            RegisterHotkeyFactory(container);

            var assembliesToRegister = AssembliesToRegister().ToArray();

            RegisterConfiguration(container, assembliesToRegister);

            RegisterParseTreeInspections(container, assembliesToRegister);
            RegisterInspections(container, assembliesToRegister);
            RegisterQuickFixes(container, assembliesToRegister);

            RegisterSpecialFactories(container);
            RegisterFactories(container, assembliesToRegister);
            
            ApplyDefaultInterfaceConvention(container, assembliesToRegister);
        }

        // note: settings namespace classes are injected in singleton scope
        private void RegisterConfiguration(IWindsorContainer container, Assembly[] assembliesToRegister)
        {
            foreach (var assembly in assembliesToRegister)
            {
                var assemblyTypes = GetIoCRegisteredTypes(assembly);

                container.Register(Classes.From(assemblyTypes)
                    .InSameNamespaceAs<Configuration>()
                    .WithService.AllInterfaces()
                    .LifestyleSingleton());
            }

            container.Register(Component.For<IPersistable<SerializableProject>>()
                .ImplementedBy<XmlPersistableDeclarations>()
                .LifestyleTransient());

            container.Register(Component.For(typeof(IPersistanceService<>), typeof(IFilePersistanceService<>))
                .ImplementedBy(typeof(XmlPersistanceService<>))
                .LifestyleSingleton());

            container.Register(Component.For<IConfigProvider<IndenterSettings>>()
                .ImplementedBy<IndenterConfigProvider>()
                .LifestyleSingleton());
            container.Register(Component.For<IConfigProvider<SourceControlSettings>>()
                .ImplementedBy<SourceControlConfigProvider>()
                .LifestyleSingleton());

            container.Register(Component.For<ISourceControlSettings>()
                .ImplementedBy<SourceControlSettings>()
                .LifestyleTransient());
        }

        private void ApplyDefaultInterfaceConvention(IWindsorContainer container, Assembly[] assembliesToRegister)
        {
            foreach (var assembly in assembliesToRegister)
            {
                var assemblyTypes = GetIoCRegisteredTypes(assembly);

                assemblyTypes.Any(t => t.Name == nameof(IMessageBox));
                
                container.Register(Classes.From(assemblyTypes)
                    .Where(type => type.Namespace != null
                                   && !type.Namespace.StartsWith("Rubberduck.VBEditor.SafeComWrappers")
                                   && !type.Name.Equals("SelectionChangeService")
                                   && !type.Name.EndsWith("Factory")
                                   && !type.Name.EndsWith("ConfigProvider")
                                   && !type.GetInterfaces().Contains(typeof(IInspection)))
                    .WithService.DefaultInterfaces()
                    .LifestyleTransient()
                );
            }
        }

        private void RegisterFactories(IWindsorContainer container, Assembly[] assembliesToRegister)
        {
            foreach (var assembly in assembliesToRegister)
            {
                var assemblyTypes = GetIoCRegisteredTypes(assembly);

                container.Register(Types.From(assemblyTypes)
                    .Where(type => type.IsInterface && type.Name.EndsWith("Factory"))
                    .Configure(c => c.AsFactory())
                    .LifestyleSingleton());
            }
        }

        private static void RegisterSpecialFactories(IWindsorContainer container)
        {
            container.Register(Component.For<IFolderBrowserFactory>()
                .ImplementedBy<DialogFactory>()
                .LifestyleSingleton());
            container.Register(Component.For<IFakesProviderFactory>()
                .ImplementedBy<FakesProviderFactory>()
                .LifestyleSingleton());
        }

        private void RegisterQuickFixes(IWindsorContainer container, Assembly[] assembliesToRegister)
        {
            foreach (var assembly in assembliesToRegister)
            {
                var assemblyTypes = GetIoCRegisteredTypes(assembly);

                container.Register(Classes.From(assemblyTypes)
                    .BasedOn<IQuickFix>()
                    .WithService.Base()
                    .WithService.Self()
                    .LifestyleSingleton());
            }
        }

        private void RegisterInspections(IWindsorContainer container, Assembly[] assembliesToRegister)
        {
            foreach (var assembly in assembliesToRegister)
            {
                var assemblyTypes = GetIoCRegisteredTypes(assembly);

                container.Register(Classes.From(assemblyTypes)
                    .BasedOn<IInspection>()
                    .WithService.Base()
                    .LifestyleTransient());
            }
        }

        private void RegisterParseTreeInspections(IWindsorContainer container, Assembly[] assembliesToRegister)
        {
            foreach (var assembly in assembliesToRegister)
            {
                var assemblyTypes = GetIoCRegisteredTypes(assembly);

                container.Register(Classes.From(assemblyTypes)
                    .BasedOn<IParseTreeInspection>()
                    .WithService.Base()
                    .WithService.Select(new[]{typeof(IInspection)})
                    .LifestyleTransient());
            }
        }

        private void RegisterRubberduckMenu(IWindsorContainer container)
        {
            const int windowMenuId = 30009;
            var controls = MainCommandBarControls(MenuBar);
            var beforeIndex = FindRubberduckMenuInsertionIndex(controls, windowMenuId);
            var menuItemTypes = RubberduckMenuItems();
            RegisterMenu<RubberduckParentMenu>(container, controls, beforeIndex, menuItemTypes);
        }

        private static void RegisterMenu<TMenu>(IWindsorContainer container, ICommandBarControls controls, int beforeIndex, Type[] menuItemTypes) where TMenu : IParentMenuItem
        {
            container.Register(Component.For<IParentMenuItem>()
                .ImplementedBy<TMenu>()
                .LifestyleTransient()
                .DependsOn(
                    Dependency.OnValue<int>(beforeIndex),
                    Dependency.OnComponentCollection<IEnumerable<IMenuItem>>(menuItemTypes))
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

        private static int FindRubberduckMenuInsertionIndex(ICommandBarControls controls, int beforeId)
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
            const int listMembersMenuId = 2529;
            var controls = MainCommandBarControls(CodeWindow);
            var beforeIndex = FindRubberduckMenuInsertionIndex(controls, listMembersMenuId);
            var menuItemTypes = CodePaneContextMenuItems();
            RegisterMenu<CodePaneContextParentMenu>(container, controls, beforeIndex, menuItemTypes);
        }

        private static Type[] CodePaneContextMenuItems()
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
            const int viewCodeMenuId = 2558;
            var controls = MainCommandBarControls(MsForms);
            var beforeIndex = FindRubberduckMenuInsertionIndex(controls, viewCodeMenuId);
            var menuItemTypes = FormDesignerContextMenuItems();
            RegisterMenu<FormDesignerContextParentMenu>(container, controls, beforeIndex, menuItemTypes);
        }

        private static Type[] FormDesignerContextMenuItems()
        {
            return new Type[]
            {
                typeof(FormDesignerRefactorRenameCommandMenuItem),
                typeof(FormDesignerFindAllReferencesCommandMenuItem)
            };
        }

        private void RegisterFormDesignerControlContextMenu(IWindsorContainer container)
        {
            const int viewCodeMenuId = 2558;
            var controls = MainCommandBarControls(MsFormsControl);
            var beforeIndex = FindRubberduckMenuInsertionIndex(controls, viewCodeMenuId);
            var menuItemTypes = FormDesignerContextMenuItems();
            RegisterMenu<FormDesignerControlContextParentMenu>(container, controls, beforeIndex, menuItemTypes);
        }

        private void RegisterProjectExplorerContextMenu(IWindsorContainer container)
        {
            const int projectPropertiesMenuId = 2578;
            var controls = MainCommandBarControls(ProjectWindow);
            var beforeIndex = FindRubberduckMenuInsertionIndex(controls, projectPropertiesMenuId);
            var menuItemTypes = ProjectWindowContextMenuItems();
            RegisterMenu<ProjectWindowContextParentMenu>(container, controls, beforeIndex, menuItemTypes);
        }

        private static Type[] ProjectWindowContextMenuItems()
        {
            return new[]
            {
                typeof(ProjectExplorerRefactorRenameCommandMenuItem),
                typeof(FindSymbolCommandMenuItem),
                typeof(FindAllReferencesCommandMenuItem),
                typeof(FindAllImplementationsCommandMenuItem)
            };
        }

        private static void RegisterRubberduckCommandBar(IWindsorContainer container)
        {
            container.Register(Component.For<RubberduckCommandBar>()
                .LifestyleTransient()
                .DependsOn(Dependency.OnComponentCollection<IEnumerable<ICommandMenuItem>>(RubberduckCommandBarItems()))
                .OnCreate((kernel, item) => item.Parent = kernel.Resolve<ICommandBars>()));
        }

        private static Type[] RubberduckCommandBarItems()
        {
            return new[]
            {
                typeof(ReparseCommandMenuItem),
                typeof(ShowParserErrorsCommandMenuItem),
                typeof(ContextSelectionLabelMenuItem),
                typeof(ContextDescriptionLabelMenuItem),
                typeof(ReferenceCounterLabelMenuItem),
#if DEBUG
                typeof(SerializeDeclarationsCommandMenuItem)
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

        private static void RegisterParentMenu<TParentMenu>(IWindsorContainer container, Type[] menuItemTypes) where TParentMenu : IParentMenuItem
        {
            container.Register(Component.For<IMenuItem, TParentMenu>()
                .ImplementedBy<TParentMenu>()
                .LifestyleTransient()
                .DependsOn(Dependency.OnComponentCollection<IEnumerable<IMenuItem>>(menuItemTypes)));
        }

        private static Type[] UnitTestingMenuItems()
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

        private static Type[] RefactoringsMenuItems()
        {
            return new[]
            {
                typeof(CodePaneRefactorRenameCommandMenuItem),
#if DEBUG
                typeof(RefactorExtractMethodCommandMenuItem),
#endif
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

        private static Type[] NavigateMenuItems()
        {
            return new[]
            {
                typeof(CodeExplorerCommandMenuItem),
#if DEBUG
                typeof(RegexSearchReplaceCommandMenuItem),
#endif
                
                typeof(FindSymbolCommandMenuItem),
                typeof(FindAllReferencesCommandMenuItem),
                typeof(FindAllImplementationsCommandMenuItem)
            };
        }

        private static Type[] SmartIndenterMenuItems()
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

            if (_initialSettings.EnableExperimentalFeatures.Any(a => a.Key == nameof(RubberduckUI.GeneralSettings_EnableSourceControl) && a.IsEnabled))
            {
                items.Add(typeof(SourceControlCommandMenuItem));
            }
            return items.ToArray();
        }

        private static void RegisterCodeExplorerViewModelWithCodeExplorerCommands(IWindsorContainer container)
        {
            var codeExplorerCommands = Assembly.GetExecutingAssembly().GetTypes()
                .Where(type => type.IsClass && type.Namespace != null &&
                               type.CustomAttributes.Any(a => a.AttributeType == typeof(CodeExplorerCommandAttribute)));
            container.Register(Component.For<CodeExplorerViewModel>()
                .DependsOn(Dependency.OnComponentCollection<List<CommandBase>>(codeExplorerCommands.ToArray()))
                .LifestyleSingleton());
        }

        private static void RegisterRefactoringDialogs(IWindsorContainer container)
        {
            container.Register(Component.For<IRefactoringDialog<RenameViewModel>>()
                .ImplementedBy<RenameDialog>()
                .LifestyleTransient());
        }

        private static void RegisterCommandMenuItems(IWindsorContainer container)
        {
            //note: The name of a registration is the full name of the implementation if not specified otherwise.
            container.Register(Classes.FromThisAssembly()
                .BasedOn<ICommandMenuItem>()
                .WithService.AllInterfaces()
                .Configure(item => item.DependsOn(Dependency.OnComponent(typeof(CommandBase),
                    CommandNameFromCommandMenuName(item.Implementation.Name))))
                .LifestyleTransient());
        }

        private static string CommandNameFromCommandMenuName(string itemName)
        {
            //note: CommandBase naming convention: [Foo]Command
            //note: ICommandMenuItem naming convention for [Foo]Command: [Foo]CommandMenuItem
            return itemName.Substring(0, itemName.Length - "MenuItem".Length);
        }

        private void RegisterCommands(IWindsorContainer container)
        {
            //note: convention: the registration name for commands is the type name, not the full type name.
            //Otherwise, namespaces would get in the way when binding to the menu items.
            RegisterCommandsWithPresenters(container);

            var commandsForCommandMenuItems = Assembly.GetExecutingAssembly().GetTypes()
                .Where(type => type.IsClass && typeof(ICommandMenuItem).IsAssignableFrom(type))
                .Select(type => CommandNameFromCommandMenuName(type.Name))
                .ToHashSet();

            container.Register(Classes.FromThisAssembly()
                .Where(type => type.Namespace != null
                            && type.Namespace.StartsWith(typeof(CommandBase).Namespace ?? string.Empty)
                            && (type.BaseType == typeof(CommandBase) || type.BaseType == typeof(RefactorCommandBase))
                            && type.Name.EndsWith("Command")
                            && commandsForCommandMenuItems.Contains(type.Name))
                .WithService.Self()
                .WithService.Select(new[] { typeof(CommandBase) })
                .LifestyleTransient()
                .Configure(c => c.Named(c.Implementation.Name)));
        }

        private void RegisterCommandsWithPresenters(IWindsorContainer container)
        {
            if (_initialSettings.EnableExperimentalFeatures.Any(a => a.Key == nameof(RubberduckUI.GeneralSettings_EnableSourceControl) && a.IsEnabled))
            {
                container.Register(Component.For<CommandBase>()
                    .ImplementedBy<SourceControlCommand>()
                    .DependsOn(Dependency.OnComponent<IDockablePresenter, SourceControlDockablePresenter>())
                    .LifestyleTransient()
                    .Named(typeof(SourceControlCommand).Name));
                container.Register(Component.For<CommandBase>()
                    .ImplementedBy<CommitCommand>()
                    .DependsOn(Dependency.OnComponent<IDockablePresenter, SourceControlDockablePresenter>())
                    .LifestyleTransient()
                    .Named(typeof(CommitCommand).Name));
                container.Register(Component.For<CommandBase>()
                    .ImplementedBy<UndoCommand>()
                    .DependsOn(Dependency.OnComponent<IDockablePresenter, SourceControlDockablePresenter>())
                    .LifestyleTransient()
                    .Named(typeof(UndoCommand).Name));
            }

            container.Register(Component.For<CommandBase>()
                .ImplementedBy<RunAllTestsCommand>()
                .DependsOn(Dependency.OnComponent<IDockablePresenter, TestExplorerDockablePresenter>())
                .LifestyleTransient()
                .Named(typeof(RunAllTestsCommand).Name));
            container.Register(Component.For<CommandBase>()
                .ImplementedBy<TestExplorerCommand>()
                .DependsOn(Dependency.OnComponent<IDockablePresenter, TestExplorerDockablePresenter>())
                .LifestyleTransient()
                .Named(typeof(TestExplorerCommand).Name));
                
            container.Register(Component.For<CommandBase>()
                .ImplementedBy<InspectionResultsCommand>()
                .DependsOn(Dependency.OnComponent<IDockablePresenter, InspectionResultsDockablePresenter>())
                .LifestyleTransient()
                .Named(typeof(InspectionResultsCommand).Name));

            container.Register(Component.For<CommandBase>()
                .ImplementedBy<CodeExplorerCommand>()
                .DependsOn(Dependency.OnComponent<IDockablePresenter, CodeExplorerDockablePresenter>())
                .LifestyleTransient()
                .Named(typeof(CodeExplorerCommand).Name));

            container.Register(Component.For<CommandBase>()
                .ImplementedBy<CodeMetricsCommand>()
                .DependsOn(Dependency.OnComponent<IDockablePresenter, CodeMetricsDockablePresenter>())
                .LifestyleSingleton()
                .Named(typeof(CodeMetricsCommand).Name));

            container.Register(Component.For<CommandBase>()
                .ImplementedBy<ToDoExplorerCommand>()
                .DependsOn(Dependency.OnComponent<IDockablePresenter, ToDoExplorerDockablePresenter>())
                .LifestyleTransient()
                .Named(typeof(ToDoExplorerCommand).Name));
        }

        private static void RegisterSmartIndenter(IWindsorContainer container)
        {
            container.Register(Component.For<IIndenter, Indenter>()
                .ImplementedBy<Indenter>()
                .LifestyleSingleton());
            container.Register(Component.For<IIndenterSettings>()
                .ImplementedBy<IndenterSettings>()
                .LifestyleSingleton());
            container.Register(Component.For<Func<IIndenterSettings>>()
                .UsingFactoryMethod(kernel => (Func<IIndenterSettings>) (() => kernel.Resolve<IGeneralConfigService>()
                    .LoadConfiguration().UserSettings.IndenterSettings))
                .LifestyleTransient()); //todo: clean up this registration
        }

        private void RegisterWindowsHooks(IWindsorContainer container)
        {
            var mainWindowHwnd = (IntPtr) _vbe.MainWindow.HWnd;
            
            container.Register(Component.For<IRubberduckHooks>()
                .ImplementedBy<RubberduckHooks>()
                .DependsOn(Dependency.OnValue<IntPtr>(mainWindowHwnd))
                .LifestyleSingleton());
        }

        private static void RegisterSourceControlControls(IWindsorContainer container)
        {
            container.Register(Component.For<SourceControlViewViewModel>()
                .LifestyleSingleton());

            container.Register(Component.For<IControlViewModel>()
                .ImplementedBy<ChangesPanelViewModel>()
                .Named(nameof(ChangesPanelViewModel))
                .LifestyleTransient());
            container.Register(Component.For<IControlView>()
                .ImplementedBy<ChangesView>()
                .DependsOn(Dependency.OnComponent<IControlViewModel, ChangesPanelViewModel>())
                .LifestyleTransient());

            container.Register(Component.For<IControlViewModel>()
                .ImplementedBy<BranchesPanelViewModel>()
                .Named(nameof(BranchesPanelViewModel))
                .LifestyleTransient());
            container.Register(Component.For<IControlView>()
                .ImplementedBy<BranchesView>()
                .DependsOn(Dependency.OnComponent<IControlViewModel, BranchesPanelViewModel>())
                .LifestyleTransient());

            container.Register(Component.For<IControlViewModel>()
                .ImplementedBy<UnsyncedCommitsPanelViewModel>()
                .Named(nameof(UnsyncedCommitsPanelViewModel))
                .LifestyleTransient());
            container.Register(Component.For<IControlView>()
                .ImplementedBy<UnsyncedCommitsView>()
                .DependsOn(Dependency.OnComponent<IControlViewModel, UnsyncedCommitsPanelViewModel>())
                .LifestyleTransient());

            container.Register(Component.For<IControlViewModel>()
                .ImplementedBy<SettingsPanelViewModel>()
                .Named(nameof(SettingsPanelViewModel))
                .LifestyleTransient());
            container.Register(Component.For<IControlView>()
                .ImplementedBy<SettingsView>()
                .DependsOn(Dependency.OnComponent<IControlViewModel, SettingsPanelViewModel>())
                .LifestyleTransient());
        }

        private static void RegisterDockableUserControls(IWindsorContainer container)
        {
            container.Register(Classes.FromThisAssembly()
                .BasedOn<IDockableUserControl>()
                .LifestyleSingleton());
        }

        private static void RegisterDockablePresenters(IWindsorContainer container)
        {
            container.Register(Component.For<IDockablePresenter>()
                .ImplementedBy<SourceControlDockablePresenter>()
                .LifestyleSingleton());
            container.Register(Component.For<IDockablePresenter>()
                .ImplementedBy<TestExplorerDockablePresenter>()
                .LifestyleSingleton());
            container.Register(Component.For<IDockablePresenter>()
                .ImplementedBy<InspectionResultsDockablePresenter>()
                .LifestyleSingleton());
            container.Register(Component.For<IDockablePresenter>()
                .ImplementedBy<CodeExplorerDockablePresenter>()
                .LifestyleSingleton());
            container.Register(Component.For<IDockablePresenter>()
                .ImplementedBy<ToDoExplorerDockablePresenter>()
                .LifestyleSingleton());
        }

        private static void SetUpCollectionResolver(IWindsorContainer container)
        {
            container.Kernel.Resolver.AddSubResolver(new CollectionResolver(container.Kernel, true));
        }

        private static void ActivateAutoMagicFactories(IWindsorContainer container)
        {
            container.Kernel.AddFacility<TypedFactoryFacility>();
        }

        private static void DeactivatePropertyInjection(IWindsorContainer container)
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
                .ImplementedBy<ProjectManager>()
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
                .Instance(() => new VBAPreprocessor(double.Parse(_vbe.Version, CultureInfo.InvariantCulture))));
        }

        private void RegisterCustomDeclarationLoadersToParser(IWindsorContainer container)
        {
            container.Register(Classes.FromAssemblyContaining<ICustomDeclarationLoader>()
                .BasedOn<ICustomDeclarationLoader>()
                .WithService.Base()
                .LifestyleSingleton());
        }

        public static IEnumerable<Assembly> AssembliesToRegister()
        {
            var assemblies = new[]
                {
                    Assembly.GetExecutingAssembly(),
                    Assembly.GetAssembly(typeof(IHostApplication)),
                    Assembly.GetAssembly(typeof(IParseCoordinator)),
                    Assembly.GetAssembly(typeof(IIndenter))
                }
                .Concat(FindPlugins());
            return assemblies;
        }

        private static IEnumerable<Assembly> FindPlugins()
        {
            var assemblies = new List<Assembly>();
            var basePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            var inspectionsAssembly = Path.Combine(basePath, "Rubberduck.Inspections.dll");
            if (File.Exists(inspectionsAssembly))
            {
                assemblies.Add(Assembly.LoadFile(inspectionsAssembly));
            }

            var path = Path.Combine(basePath, "Plug-ins");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            foreach (var library in Directory.EnumerateFiles(path, "*.dll"))
            {
                try
                {
                    assemblies.Add(Assembly.LoadFile(library));
                }
                catch (Exception)
                {
                    // can we log yet?
                }
            }

            return assemblies;
        }

        private static void RegisterAppWithSpecialDependencies(IWindsorContainer container)
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
            container.Register(Component.For<ICommandBars>().Instance(_vbe.CommandBars)); //note: This registration makes Castle Windsor inject _vbe_CommandBars in all ICommandBars Parent properties.
        }

        private static void RegisterHotkeyFactory(IWindsorContainer container)
        {
            container.Register(Component.For<HotkeyFactory>().LifestyleSingleton());
        }

        private IEnumerable<TypeInfo> GetIoCRegisteredTypes(Assembly assembly)
        {
            return assembly.DefinedTypes
                .Where(w =>
                {
                    var attribute = w.CustomAttributes.FirstOrDefault(f => f.AttributeType == typeof(ExperimentalAttribute));
                    var ctorArg = attribute?.ConstructorArguments.Any() == true ? (string)attribute.ConstructorArguments.First().Value : string.Empty;

                    return attribute == null ||
                           _initialSettings.EnableExperimentalFeatures.Any(a => a.Key == ctorArg && a.IsEnabled);
                });
        }
    }
}
