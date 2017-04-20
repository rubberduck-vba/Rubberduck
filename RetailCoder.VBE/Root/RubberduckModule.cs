using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Ninject;
using Ninject.Extensions.Conventions;
using Ninject.Modules;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols.DeclarationLoaders;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;
using Rubberduck.SourceControl;
using Rubberduck.UI;
using Rubberduck.UI.CodeExplorer;
using Rubberduck.UI.Inspections;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;
using Rubberduck.UI.Controls;
using Rubberduck.UI.SourceControl;
using Rubberduck.UI.ToDoItems;
using Rubberduck.UI.UnitTesting;
using System.Globalization;
using System.IO;
using Ninject.Extensions.Interception.Infrastructure.Language;
using Ninject.Extensions.NamedScope;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.UI.Command.MenuItems.CommandBars;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;
using ReparseCommandMenuItem = Rubberduck.UI.Command.MenuItems.CommandBars.ReparseCommandMenuItem;
using Rubberduck.UI.Refactorings;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.UnitTesting;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.PreProcessing;

namespace Rubberduck.Root
{
    public class RubberduckModule : NinjectModule
    {
        private readonly IVBE _vbe;
        private readonly IAddIn _addin;
        
        private const int MenuBar = 1;
        private const int CodeWindow = 9;
        private const int ProjectWindow = 14;
        private const int MsForms = 17;
        private const int MsFormsControl = 18;

        public RubberduckModule(IVBE vbe, IAddIn addin)
        {
            _vbe = vbe;
            _addin = addin;
        }

        public override void Load()
        {
            // bind VBE and AddIn dependencies to host-provided instances.
            Bind<IVBE>().ToConstant(_vbe);
            Bind<IAddIn>().ToConstant(_addin);
            Bind<App>().ToSelf().InSingletonScope();
            Bind<RubberduckParserState, IParseTreeProvider, IDeclarationFinderProvider>().To<RubberduckParserState>().InSingletonScope();
            Bind<ISelectionChangeService>().To<SelectionChangeService>().InSingletonScope();
            Bind<ISourceControlProvider>().To<GitProvider>();
            //Bind<GitProvider>().ToSelf().InSingletonScope();        
            Bind<TestExplorerModel>().ToSelf().InSingletonScope();
            Bind<IOperatingSystem>().To<WindowsOperatingSystem>().InSingletonScope();
            
            Bind<CommandBase>().To<VersionCheckCommand>().WhenInjectedExactlyInto<App>();

            var assemblies = new[]
            {
                Assembly.GetExecutingAssembly(),
                Assembly.GetAssembly(typeof(IHostApplication)),
                Assembly.GetAssembly(typeof(IParseCoordinator)),
                Assembly.GetAssembly(typeof(IIndenter))
            }
            .Concat(FindPlugins())
            .ToArray();

            BindCodeInspectionTypes(assemblies);
            BindQuickFixTypes(assemblies);
            BindRefactoringDialogs();

            ApplyDefaultInterfacesConvention(assemblies);
            ApplyConfigurationConvention(assemblies);
            ApplyAbstractFactoryConvention(assemblies);
            Rebind<IFolderBrowserFactory>().To<DialogFactory>().InSingletonScope();
            Rebind<IFakesProviderFactory>().To<FakesProviderFactory>().InSingletonScope();

            BindCommandsToMenuItems();

            Rebind<IIndenter>().To<Indenter>().InSingletonScope();            
            Rebind<IIndenterSettings>().To<IndenterSettings>();
            Bind<Func<IIndenterSettings>>().ToMethod(t => () => KernelInstance.Get<IGeneralConfigService>().LoadConfiguration().UserSettings.IndenterSettings);

            BindCustomDeclarationLoadersToParser();
            Rebind<ICOMReferenceSynchronizer, IProjectReferencesProvider>().To<COMReferenceSynchronizer>().InSingletonScope().WithConstructorArgument("serializedDeclarationsPath", (string)null);
            Rebind<IBuiltInDeclarationLoader>().To<BuiltInDeclarationLoader>().InSingletonScope();
            Rebind<IDeclarationResolveRunner>().To<DeclarationResolveRunner>().InSingletonScope();
            Rebind<IModuleToModuleReferenceManager>().To<ModuleToModuleReferenceManager>().InSingletonScope();
            Rebind<IParserStateManager>().To<ParserStateManager>().InSingletonScope();
            Rebind<IParseRunner>().To<ParseRunner>().InSingletonScope();
            Rebind<IParsingStageService>().To<ParsingStageService>().InSingletonScope();
            Rebind<IProjectManager>().To<ProjectManager>().InSingletonScope();
            Rebind<IReferenceRemover>().To<ReferenceRemover>().InSingletonScope();
            Rebind<IReferenceResolveRunner>().To<ReferenceResolveRunner>().InSingletonScope();
            Rebind<IParseCoordinator>().To<ParseCoordinator>().InSingletonScope();

            Bind<Func<IVBAPreprocessor>>().ToMethod(p => () => new VBAPreprocessor(double.Parse(_vbe.Version, CultureInfo.InvariantCulture)));

            Rebind<ISearchResultsWindowViewModel>().To<SearchResultsWindowViewModel>().InSingletonScope();

            Bind<SourceControlViewViewModel>().ToSelf().InSingletonScope();
            Bind<IControlView>().To<ChangesView>().InCallScope();
            Bind<IControlView>().To<BranchesView>().InCallScope();
            Bind<IControlView>().To<UnsyncedCommitsView>().InCallScope();
            Bind<IControlView>().To<SettingsView>().InCallScope();

            Bind<IControlViewModel>().To<ChangesViewViewModel>().WhenInjectedInto<ChangesView>().InCallScope();
            Bind<IControlViewModel>().To<BranchesViewViewModel>().WhenInjectedInto<BranchesView>().InCallScope();
            Bind<IControlViewModel>().To<UnsyncedCommitsViewViewModel>().WhenInjectedInto<UnsyncedCommitsView>().InCallScope();
            Bind<IControlViewModel>().To<SettingsViewViewModel>().WhenInjectedInto<SettingsView>().InCallScope();

            Bind<SearchResultPresenterInstanceManager>()
                .ToSelf()
                .InSingletonScope();

            Bind<IDockablePresenter>().To<SourceControlDockablePresenter>()
                .WhenInjectedInto(
                    typeof(ShowSourceControlPanelCommand),
                    typeof(CommitCommand),
                    typeof(UndoCommand))
                .InSingletonScope();

            Bind<IDockablePresenter>().To<TestExplorerDockablePresenter>()
                .WhenInjectedInto(
                    typeof (RunAllTestsCommand), 
                    typeof (TestExplorerCommand))
                .InSingletonScope();

            Bind<IDockablePresenter>().To<CodeInspectionsDockablePresenter>()
                .WhenInjectedInto<InspectionResultsCommand>()
                .InSingletonScope();

            Bind<IDockablePresenter>().To<CodeExplorerDockablePresenter>()
                .WhenInjectedInto<CodeExplorerCommand>()
                .InSingletonScope();

            Bind<IDockablePresenter>().To<ToDoExplorerDockablePresenter>()
                .WhenInjectedInto<ToDoExplorerCommand>()
                .InSingletonScope();

            BindDockableToolwindows();
            BindCommandsToCodeExplorer();
            ConfigureRubberduckCommandBar();
            ConfigureRubberduckMenu();
            ConfigureCodePaneContextMenu();
            ConfigureFormDesignerContextMenu();
            ConfigureFormDesignerControlContextMenu();
            ConfigureProjectExplorerContextMenu();
            
            BindWindowsHooks();
        }

        private IEnumerable<Assembly> FindPlugins()
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

        private void BindRefactoringDialogs()
        {
            Bind<IRefactoringDialog<RenameViewModel>>().To<RenameDialog>();
        }

        private void BindDockableToolwindows()
        {
            Kernel.Bind(t => t.FromThisAssembly()
                .SelectAllClasses()
                .InheritedFrom<IDockableUserControl>()
                .BindToSelf()
                .Configure(binding => binding.InSingletonScope()));
        }

        private void BindWindowsHooks()
        {
            Rebind<IAttachable>().To<TimerHook>()
                .InSingletonScope()
                .WithConstructorArgument("mainWindowHandle", (IntPtr)_vbe.MainWindow.HWnd);

            Rebind<IRubberduckHooks>().To<RubberduckHooks>()
                .InSingletonScope()
                .WithConstructorArgument("mainWindowHandle", (IntPtr)_vbe.MainWindow.HWnd);
        }

        private void ApplyDefaultInterfacesConvention(IEnumerable<Assembly> assemblies)
        {
            Kernel.Bind(t => t.From(assemblies)
                .SelectAllClasses()
                // inspections & factories have their own binding rules
                .Where(type => type.Namespace != null
                            && !type.Namespace.StartsWith("Rubberduck.VBEditor.SafeComWrappers")
                            && !type.Name.Equals("SelectionChangeService")
                            && !type.Name.EndsWith("Factory") && !type.Name.EndsWith("ConfigProvider") && !type.GetInterfaces().Contains(typeof(IInspection)))
                .BindDefaultInterface()
                .Configure(binding => binding.InCallScope())); // TransientScope wouldn't dispose disposables
        }

        // note: settings namespace classes are injected in singleton scope
        private void ApplyConfigurationConvention(IEnumerable<Assembly> assemblies)
        {
            Kernel.Bind(t => t.From(assemblies)
                .SelectAllClasses()
                .InNamespaceOf<Configuration>()
                .BindAllInterfaces()
                .Configure(binding => binding.InSingletonScope()));

            Bind<IPersistable<SerializableProject>>().To<XmlPersistableDeclarations>().InCallScope();

            Bind<IPersistanceService<CodeInspectionSettings>>().To<XmlPersistanceService<CodeInspectionSettings>>().InSingletonScope();
            Bind<IPersistanceService<GeneralSettings>>().To<XmlPersistanceService<GeneralSettings>>().InSingletonScope();
            Bind<IPersistanceService<HotkeySettings>>().To<XmlPersistanceService<HotkeySettings>>().InSingletonScope();
            Bind<IPersistanceService<ToDoListSettings>>().To<XmlPersistanceService<ToDoListSettings>>().InSingletonScope();
            Bind<IPersistanceService<UnitTestSettings>>().To<XmlPersistanceService<UnitTestSettings>>().InSingletonScope();
            Bind<IPersistanceService<WindowSettings>>().To<XmlPersistanceService<WindowSettings>>().InSingletonScope();
            Bind<IPersistanceService<IndenterSettings>>().To<XmlPersistanceService<IndenterSettings>>().InSingletonScope();
            Bind<IFilePersistanceService<SourceControlSettings>>().To<XmlPersistanceService<SourceControlSettings>>().InSingletonScope();

            Bind<IConfigProvider<IndenterSettings>>().To<IndenterConfigProvider>().InSingletonScope();
            Bind<IConfigProvider<SourceControlSettings>>().To<SourceControlConfigProvider>().InSingletonScope();

            Bind<ICodeInspectionSettings>().To<CodeInspectionSettings>().InCallScope();
            Bind<IGeneralSettings>().To<GeneralSettings>().InCallScope();
            Bind<IHotkeySettings>().To<HotkeySettings>().InCallScope();
            Bind<IToDoListSettings>().To<ToDoListSettings>().InCallScope();
            Bind<IUnitTestSettings>().To<UnitTestSettings>().InCallScope();
            Bind<IWindowSettings>().To<WindowSettings>().InCallScope();
            Bind<IIndenterSettings>().To<IndenterSettings>().InCallScope();
            Bind<ISourceControlSettings>().To<SourceControlSettings>().InCallScope();        
        }

        // note convention: abstract factory interface names end with "Factory".
        private void ApplyAbstractFactoryConvention(IEnumerable<Assembly> assemblies)
        {
            Kernel.Bind(t => t.From(assemblies)
                .SelectAllInterfaces()
                .Where(type => type.Name.EndsWith("Factory")) 
                .BindToFactory()
                .Configure(binding => binding.InSingletonScope()));
        }

        // note: IInspection implementations are discovered in all assemblies, via reflection.
        private void BindCodeInspectionTypes(IEnumerable<Assembly> assemblies)
        {
            var inspections = assemblies
                .SelectMany(a => a.GetTypes().Where(type => type.IsClass && !type.IsAbstract && type.GetInterfaces().Contains(typeof (IInspection))))
                .ToList();

            // multibinding for IEnumerable<IInspection> dependency
            foreach (var inspection in inspections)
            {
                var iParseTreeInspection = inspection.GetInterfaces().SingleOrDefault(i => i.Name == "IParseTreeInspection");
                if (iParseTreeInspection != null)
                {
                    var binding = Bind(iParseTreeInspection)
                        .To(inspection)
                        .InCallScope()
                        .Named(inspection.FullName);

                    binding.Intercept().With<TimedCallLoggerInterceptor>();
                    binding.Intercept().With<EnumerableCounterInterceptor<IInspectionResult>>();

                    var localInspection = inspection;
                    Bind<IInspection>().ToMethod(c => (IInspection)c.Kernel.Get(iParseTreeInspection, localInspection.FullName));
                }
                else
                {
                    var binding = Bind<IInspection>().To(inspection).InCallScope();
                    binding.Intercept().With<TimedCallLoggerInterceptor>();
                    binding.Intercept().With<EnumerableCounterInterceptor<IInspectionResult>>();
                }
            }
        }

        // note: IQuickFix implementations are discovered in all assemblies, via reflection.
        private void BindQuickFixTypes(IEnumerable<Assembly> assemblies)
        {
            var quickFixes = assemblies
                .SelectMany(a => a.GetTypes().Where(type => type.IsClass && !type.IsAbstract && type.GetInterfaces().Contains(typeof(IQuickFix))))
                .ToList();

            // multibinding for IEnumerable<IQuickFix> dependency
            foreach (var quickFix in quickFixes)
            {
                Bind<IQuickFix>().To(quickFix).InSingletonScope();
            }
        }

        private void ConfigureRubberduckCommandBar()
        {
            var commandBars = _vbe.CommandBars;
            var items = GetRubberduckCommandBarItems();
            Bind<RubberduckCommandBar>()
                .ToSelf()
                .InCallScope()
                .WithPropertyValue("Parent", commandBars)
                .WithConstructorArgument("items", items);
        }

        private void ConfigureRubberduckMenu()
        {
            const int windowMenuId = 30009;
            var commandBars = _vbe.CommandBars;
            var menuBar = commandBars[MenuBar];
            var controls = menuBar.Controls;
            var beforeIndex = FindRubberduckMenuInsertionIndex(controls, windowMenuId);
            var items = GetRubberduckMenuItems();
            BindParentMenuItem<RubberduckParentMenu>(controls, beforeIndex, items);
        }

        private void ConfigureCodePaneContextMenu()
        {
            const int listMembersMenuId = 2529;
            var commandBars = _vbe.CommandBars;
            var menuBar = commandBars[CodeWindow];
            var controls = menuBar.Controls;
            var beforeControl = controls.FirstOrDefault(control => control.Id == listMembersMenuId);
            var beforeIndex = beforeControl == null ? 1 : beforeControl.Index;
            var items = GetCodePaneContextMenuItems();
            BindParentMenuItem<CodePaneContextParentMenu>(controls, beforeIndex, items);
        }

        private void ConfigureFormDesignerContextMenu()
        {
            const int viewCodeMenuId = 2558;
            var commandBars = _vbe.CommandBars;
            var menuBar = commandBars[MsForms];
            var controls = menuBar.Controls;
            var beforeControl = controls.FirstOrDefault(control => control.Id == viewCodeMenuId);
            var beforeIndex = beforeControl == null ? 1 : beforeControl.Index;
            var items = GetFormDesignerContextMenuItems();
            BindParentMenuItem<FormDesignerContextParentMenu>(controls, beforeIndex, items);
        }

        private void ConfigureFormDesignerControlContextMenu()
        {
            const int viewCodeMenuId = 2558;
            var commandBars = _vbe.CommandBars;
            var menuBar = commandBars[MsFormsControl];
            var controls = menuBar.Controls;
            var beforeControl = controls.FirstOrDefault(control => control.Id == viewCodeMenuId);
            var beforeIndex = beforeControl == null ? 1 : beforeControl.Index;
            var items = GetFormDesignerContextMenuItems();
            BindParentMenuItem<FormDesignerControlContextParentMenu>(controls, beforeIndex, items);
        }

        private void ConfigureProjectExplorerContextMenu()
        {
            const int projectPropertiesMenuId = 2578;
            var commandBars = _vbe.CommandBars;
            var menuBar = commandBars[ProjectWindow];
            var controls = menuBar.Controls;
            var beforeControl = controls.FirstOrDefault(control => control.Id == projectPropertiesMenuId);
            var beforeIndex = beforeControl == null ? 1 : beforeControl.Index;
            var items = GetProjectWindowContextMenuItems();
            BindParentMenuItem<ProjectWindowContextParentMenu>(controls, beforeIndex, items);
        }

        private void BindParentMenuItem<TParentMenu>(ICommandBarControls parent, int beforeIndex, IEnumerable<IMenuItem> items)
        {
            Bind<IParentMenuItem>().To(typeof(TParentMenu))
                .WhenInjectedInto<IAppMenu>()
                .InCallScope()
                .WithConstructorArgument("items", items)
                .WithConstructorArgument("beforeIndex", beforeIndex)
                .WithPropertyValue("Parent", parent);
        }

        private static int FindRubberduckMenuInsertionIndex(ICommandBarControls controls, int beforeId)
        {
            for (var i = 1; i <= controls.Count; i++)
            {
                var item = controls[i];
                if (item.IsBuiltIn && item.Id == beforeId)
                {
                    return i;
                }
            }

            return controls.Count;
        }

        private void BindCommandsToMenuItems()
        {
            var types = Assembly.GetExecutingAssembly().GetTypes()
                .Where(type => type.IsClass && type.Namespace != null && type.Namespace.StartsWith(typeof(CommandBase).Namespace ?? string.Empty))
                .ToList();

            // note: CommandBase naming convention: [Foo]Command
            var baseCommandTypes = new[] { typeof(CommandBase), typeof(RefactorCommandBase) };
            var commands = types.Where(type => type.IsClass && baseCommandTypes.Contains(type.BaseType) && type.Name.EndsWith("Command"));
            foreach (var command in commands)
            {
                var commandName = command.Name.Substring(0, command.Name.Length - "Command".Length);
                try
                {
                    // note: ICommandMenuItem naming convention for [Foo]Command: [Foo]CommandMenuItem
                    var item = types.SingleOrDefault(type => type.Name == commandName + "CommandMenuItem");
                    if (item != null)
                    {
                        var binding = Bind<CommandBase>().To(command);
                        var whenCommandMenuItemCondition =
                            binding.WhenInjectedInto(item).BindingConfiguration.Condition;
                        var whenHooksCondition =
                            binding.WhenInjectedInto<RubberduckHooks>().BindingConfiguration.Condition;

                        binding.When(request => whenCommandMenuItemCondition(request) || whenHooksCondition(request))
                            .InCallScope();
                    }
                }
                catch (InvalidOperationException)
                {
                    // rename one of the classes, "FooCommand" is expected to match exactly 1 "FooBarXyzCommandMenuItem"
                }
            }
        }

        private void BindCommandsToCodeExplorer()
        {
            var commands = Assembly.GetExecutingAssembly().GetTypes()
                .Where(type => type.IsClass && type.Namespace != null &&
                               type.CustomAttributes.Any(a => a.AttributeType == typeof(CodeExplorerCommandAttribute)));

            foreach (var command in commands)
            {
                Bind<CommandBase>().To(command).InSingletonScope();
            }
        }

        private void BindCustomDeclarationLoadersToParser()
        {
            var loaders = Assembly.GetAssembly(typeof(ICustomDeclarationLoader))
                          .GetTypes()
                          .Where(type => type.GetInterfaces().Contains(typeof(ICustomDeclarationLoader)));

            foreach (var loader in loaders)
            {
                Bind<ICustomDeclarationLoader>().To(loader).InSingletonScope();
            }
        }

        private IEnumerable<ICommandMenuItem> GetRubberduckCommandBarItems()
        {
            return new ICommandMenuItem[]
            {
                KernelInstance.Get<ReparseCommandMenuItem>(),
                KernelInstance.Get<ShowParserErrorsCommandMenuItem>(),
                KernelInstance.Get<ContextSelectionLabelMenuItem>(),
                KernelInstance.Get<ContextDescriptionLabelMenuItem>(),
                KernelInstance.Get<ReferenceCounterLabelMenuItem>(),
#if DEBUG
                KernelInstance.Get<SerializeDeclarationsCommandMenuItem>()
#endif
            };
        }

        private IEnumerable<IMenuItem> GetRubberduckMenuItems()
        {
            return new[]
            {
                KernelInstance.Get<AboutCommandMenuItem>(),
                KernelInstance.Get<SettingsCommandMenuItem>(),
                KernelInstance.Get<InspectionResultsCommandMenuItem>(),
                GetUnitTestingParentMenu(),
                GetSmartIndenterParentMenu(),
                GetToolsParentMenu(),
                GetRefactoringsParentMenu(),
                GetNavigateParentMenu()
            };
        }

        private IMenuItem GetUnitTestingParentMenu()
        {
            var items = new IMenuItem[]
            {
                KernelInstance.Get<RunAllTestsCommandMenuItem>(),
                KernelInstance.Get<TestExplorerCommandMenuItem>(),
                KernelInstance.Get<AddTestModuleCommandMenuItem>(),
                KernelInstance.Get<AddTestMethodCommandMenuItem>(),
                KernelInstance.Get<AddTestMethodExpectedErrorCommandMenuItem>()
            };
            return new UnitTestingParentMenu(items);
        }

        private IMenuItem GetRefactoringsParentMenu()
        {
            var items = new IMenuItem[]
            {
                KernelInstance.Get<CodePaneRefactorRenameCommandMenuItem>(),
#if DEBUG
                KernelInstance.Get<RefactorExtractMethodCommandMenuItem>(),
#endif
                KernelInstance.Get<RefactorReorderParametersCommandMenuItem>(),
                KernelInstance.Get<RefactorRemoveParametersCommandMenuItem>(),
                KernelInstance.Get<RefactorIntroduceParameterCommandMenuItem>(),
                KernelInstance.Get<RefactorIntroduceFieldCommandMenuItem>(),
                KernelInstance.Get<RefactorEncapsulateFieldCommandMenuItem>(),
                KernelInstance.Get<RefactorMoveCloserToUsageCommandMenuItem>(),
                KernelInstance.Get<RefactorExtractInterfaceCommandMenuItem>(),
                KernelInstance.Get<RefactorImplementInterfaceCommandMenuItem>()
            };
            return new RefactoringsParentMenu(items);
        }

        private IMenuItem GetNavigateParentMenu()
        {
            var items = new IMenuItem[]
            {
                KernelInstance.Get<CodeExplorerCommandMenuItem>(),
                //KernelInstance.Get<RegexSearchReplaceCommandMenuItem>(),
                KernelInstance.Get<FindSymbolCommandMenuItem>(),
                KernelInstance.Get<FindAllReferencesCommandMenuItem>(),
                KernelInstance.Get<FindAllImplementationsCommandMenuItem>()
            };
            return new NavigateParentMenu(items);
        }

        private IMenuItem GetSmartIndenterParentMenu()
        {
            var items = new IMenuItem[]
            {
                KernelInstance.Get<IndentCurrentProcedureCommandMenuItem>(),
                KernelInstance.Get<IndentCurrentModuleCommandMenuItem>(),
                KernelInstance.Get<IndentCurrentProjectCommandMenuItem>(),
                KernelInstance.Get<NoIndentAnnotationCommandMenuItem>()
            };

            return new SmartIndenterParentMenu(items);
        }

        private IEnumerable<IMenuItem> GetCodePaneContextMenuItems()
        {
            return new[]
            {
                GetRefactoringsParentMenu(),
                GetSmartIndenterParentMenu(),
                KernelInstance.Get<FindSymbolCommandMenuItem>(),
                KernelInstance.Get<FindAllReferencesCommandMenuItem>(),
                KernelInstance.Get<FindAllImplementationsCommandMenuItem>()
            };
        }

        private IMenuItem GetToolsParentMenu()
        {
            var items = new IMenuItem[]
            {
                KernelInstance.Get<ShowSourceControlPanelCommandMenuItem>(),
                KernelInstance.Get<RegexAssistantCommandMenuItem>(),
                KernelInstance.Get<ToDoExplorerCommandMenuItem>()
            };

            return new ToolsParentMenu(items);
        }

        private IEnumerable<IMenuItem> GetFormDesignerContextMenuItems()
        {
            return new IMenuItem[]
            {
                KernelInstance.Get<FormDesignerRefactorRenameCommandMenuItem>()
            };
        }

        private IEnumerable<IMenuItem> GetProjectWindowContextMenuItems()
        {
            return new IMenuItem[]
            {
                KernelInstance.Get<ProjectExplorerRefactorRenameCommandMenuItem>(),
                KernelInstance.Get<FindSymbolCommandMenuItem>(),
                KernelInstance.Get<FindAllReferencesCommandMenuItem>(),
                KernelInstance.Get<FindAllImplementationsCommandMenuItem>()
            };
        }
    }
}
