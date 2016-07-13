using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Ninject;
using Ninject.Extensions.Conventions;
using Ninject.Modules;
using Rubberduck.Common;
using Rubberduck.Inspections;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing;
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
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.VBEHost;
using Rubberduck.Parsing.Preprocessing;
using System.Globalization;
using Ninject.Extensions.Interception.Infrastructure.Language;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.CodeExplorer.Commands;

namespace Rubberduck.Root
{
    public class RubberduckModule : NinjectModule
    {
        private readonly VBE _vbe;
        private readonly AddIn _addin;

        private const int MenuBar = 1;
        private const int CodeWindow = 9;
        private const int ProjectWindow = 14;
        private const int MsForms = 17;
        private const int MsFormsControl = 18;

        public RubberduckModule(VBE vbe, AddIn addin)
        {
            _vbe = vbe;
            _addin = addin;
        }

        public override void Load()
        {
            // bind VBE and AddIn dependencies to host-provided instances.
            Bind<VBE>().ToConstant(_vbe);
            Bind<AddIn>().ToConstant(_addin);
            Bind<Sinks>().ToSelf().InSingletonScope();
            Bind<App>().ToSelf().InSingletonScope();
            Bind<RubberduckParserState>().ToSelf().InSingletonScope();
            Bind<GitProvider>().ToSelf().InSingletonScope();
            Bind<NewUnitTestModuleCommand>().ToSelf().InSingletonScope();
            Bind<NewTestMethodCommand>().ToSelf().InSingletonScope();
            Bind<RubberduckCommandBar>().ToSelf().InSingletonScope();
            Bind<TestExplorerModel>().ToSelf().InSingletonScope();
            Bind<IOperatingSystem>().To<WindowsOperatingSystem>().InSingletonScope();

            BindCodeInspectionTypes();

            var assemblies = new[]
            {
                Assembly.GetExecutingAssembly(),
                Assembly.GetAssembly(typeof(IHostApplication)),
                Assembly.GetAssembly(typeof(IRubberduckParser)),
                Assembly.GetAssembly(typeof(IIndenter))
            };

            ApplyConfigurationConvention(assemblies);
            ApplyDefaultInterfacesConvention(assemblies);
            ApplyAbstractFactoryConvention(assemblies);

            BindCommandsToMenuItems();

            Rebind<Rubberduck.Parsing.ISinks>().To<Sinks>().InSingletonScope();
            Rebind<IIndenter>().To<Indenter>().InSingletonScope();
            Rebind<IIndenterSettings>().To<IndenterSettings>();
            Bind<Func<IIndenterSettings>>().ToMethod(t => () => Kernel.Get<IGeneralConfigService>().LoadConfiguration().UserSettings.IndenterSettings);

            BindCustomDeclarationLoadersToParser();
            Rebind<IRubberduckParser>().To<RubberduckParser>().InSingletonScope();
            Bind<Func<IVBAPreprocessor>>().ToMethod(p => () => new VBAPreprocessor(double.Parse(_vbe.Version, CultureInfo.InvariantCulture)));

            Rebind<ISearchResultsWindowViewModel>().To<SearchResultsWindowViewModel>().InSingletonScope();
            Bind<SearchResultPresenterInstanceManager>().ToSelf().InSingletonScope();

            Bind<IPresenter>().To<TestExplorerDockablePresenter>()
                .WhenInjectedInto<TestExplorerCommand>()
                .InSingletonScope()
                .WithConstructorArgument<IDockableUserControl>(new TestExplorerWindow { ViewModel = Kernel.Get<TestExplorerViewModel>() });

            Bind<IPresenter>().To<CodeInspectionsDockablePresenter>()
                .WhenInjectedInto<InspectionResultsCommand>()
                .InSingletonScope()
                .WithConstructorArgument<IDockableUserControl>(new CodeInspectionsWindow { ViewModel = Kernel.Get<InspectionResultsViewModel>() });

            Bind<IControlView>().To<ChangesView>().Named("changesView");
            Bind<IControlView>().To<BranchesView>().Named("branchesView");
            Bind<IControlView>().To<UnsyncedCommitsView>().Named("unsyncedCommitsView");
            Bind<IControlView>().To<SettingsView>().Named("settingsView");

            Bind<IControlViewModel>().To<ChangesViewViewModel>()
                .WhenInjectedInto<ChangesView>();
            Bind<IControlViewModel>().To<BranchesViewViewModel>()
                .WhenInjectedInto<BranchesView>();
            Bind<IControlViewModel>().To<UnsyncedCommitsViewViewModel>()
                .WhenInjectedInto<UnsyncedCommitsView>();
            Bind<IControlViewModel>().To<SettingsViewViewModel>()
                .WhenInjectedInto<SettingsView>();

            Bind<ISourceControlProviderFactory>().To<SourceControlProviderFactory>()
                .WhenInjectedInto<SourceControlViewViewModel>();

            Bind<SourceControlDockablePresenter>().ToSelf()
                .InSingletonScope()
                .WithConstructorArgument(new SourceControlPanel { ViewModel = Kernel.Get<SourceControlViewViewModel>() });
            
            BindCommandsToCodeExplorer();
            Bind<IPresenter>().To<CodeExplorerDockablePresenter>()
                .WhenInjectedInto<CodeExplorerCommand>()
                .InSingletonScope()
                .WithConstructorArgument<IDockableUserControl>(new CodeExplorerWindow { ViewModel = Kernel.Get<CodeExplorerViewModel>() });

            Bind<IPresenter>().To<ToDoExplorerDockablePresenter>()
                .WhenInjectedInto<ToDoExplorerCommand>()
                .InSingletonScope()
                .WithConstructorArgument<IDockableUserControl>(new ToDoExplorerWindow { ViewModel = Kernel.Get<ToDoExplorerViewModel>() });

            ConfigureRubberduckMenu();
            ConfigureCodePaneContextMenu();
            ConfigureFormDesignerContextMenu();
            ConfigureFormDesignerControlContextMenu();
            ConfigureProjectExplorerContextMenu();
            

            BindWindowsHooks();
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
                .Where(type => !type.Name.EndsWith("Factory") && !type.Name.EndsWith("ConfigProvider") && !type.GetInterfaces().Contains(typeof(IInspection)))
                .BindDefaultInterface()
                .Configure(binding => binding.InThreadScope())); // TransientScope wouldn't dispose disposables
        }

        // note: settings namespace classes are injected in singleton scope
        private void ApplyConfigurationConvention(IEnumerable<Assembly> assemblies)
        {
            Kernel.Bind(t => t.From(assemblies)
                .SelectAllClasses()
                .InNamespaceOf<Configuration>()
                .BindAllInterfaces()
                .Configure(binding => binding.InSingletonScope()));

            Bind<IPersistanceService<CodeInspectionSettings>>().To<XmlPersistanceService<CodeInspectionSettings>>().InSingletonScope();
            Bind<IPersistanceService<GeneralSettings>>().To<XmlPersistanceService<GeneralSettings>>().InSingletonScope();
            Bind<IPersistanceService<HotkeySettings>>().To<XmlPersistanceService<HotkeySettings>>().InSingletonScope();
            Bind<IPersistanceService<ToDoListSettings>>().To<XmlPersistanceService<ToDoListSettings>>().InSingletonScope();
            Bind<IPersistanceService<UnitTestSettings>>().To<XmlPersistanceService<UnitTestSettings>>().InSingletonScope();
            Bind<IPersistanceService<IndenterSettings>>().To<XmlPersistanceService<IndenterSettings>>().InSingletonScope();
            Bind<IFilePersistanceService<SourceControlSettings>>().To<XmlPersistanceService<SourceControlSettings>>().InSingletonScope();

            Bind<IIndenterConfigProvider>().To<IndenterConfigProvider>().InSingletonScope();
            Bind<ISourceControlConfigProvider>().To<SourceControlConfigProvider>().InSingletonScope();

            Bind<ICodeInspectionSettings>().To<CodeInspectionSettings>();
            Bind<IGeneralSettings>().To<GeneralSettings>();
            Bind<IHotkeySettings>().To<HotkeySettings>();
            Bind<IToDoListSettings>().To<ToDoListSettings>();
            Bind<IUnitTestSettings>().To<UnitTestSettings>();
            Bind<IIndenterSettings>().To<IndenterSettings>();
            Bind<ISourceControlSettings>().To<SourceControlSettings>();        
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

        // note: InspectionBase implementations are discovered in the Rubberduck assembly via reflection.
        private void BindCodeInspectionTypes()
        {
            var inspections = Assembly.GetExecutingAssembly()
                                      .GetTypes()
                                      .Where(type => type.BaseType == typeof (InspectionBase));

            // multibinding for IEnumerable<IInspection> dependency
            foreach (var inspection in inspections)
            {
                if (typeof(IParseTreeInspection).IsAssignableFrom(inspection))
                {
                    var binding = Bind<IParseTreeInspection>()
                        .To(inspection)
                        .InSingletonScope()
                        .Named(inspection.FullName);

                    binding.Intercept().With<TimedCallLoggerInterceptor>();
                    binding.Intercept().With<EnumerableCounterInterceptor<InspectionResultBase>>();

                    Bind<IInspection>().ToMethod(
                        c => c.Kernel.Get<IParseTreeInspection>(inspection.FullName));
                }
                else
                {
                    var binding = Bind<IInspection>().To(inspection).InSingletonScope();
                    binding.Intercept().With<TimedCallLoggerInterceptor>();
                    binding.Intercept().With<EnumerableCounterInterceptor<InspectionResultBase>>();
                }
            }
        }

        private void ConfigureRubberduckMenu()
        {
            const int windowMenuId = 30009;
            var parent = Kernel.Get<VBE>().CommandBars[MenuBar].Controls;
            var beforeIndex = FindRubberduckMenuInsertionIndex(parent, windowMenuId);

            var items = GetRubberduckMenuItems();
            BindParentMenuItem<RubberduckParentMenu>(parent, beforeIndex, items);
        }

        private void ConfigureCodePaneContextMenu()
        {
            const int listMembersMenuId = 2529;
            var parent = Kernel.Get<VBE>().CommandBars[CodeWindow].Controls;
            var beforeControl = parent.Cast<CommandBarControl>().FirstOrDefault(control => control.Id == listMembersMenuId);
            var beforeIndex = beforeControl == null ? 1 : beforeControl.Index;

            var items = GetCodePaneContextMenuItems();
            BindParentMenuItem<CodePaneContextParentMenu>(parent, beforeIndex, items);
        }

        private void ConfigureFormDesignerContextMenu()
        {
            const int viewCodeMenuId = 2558;
            var parent = Kernel.Get<VBE>().CommandBars[MsForms].Controls;
            var beforeControl = parent.Cast<CommandBarControl>().FirstOrDefault(control => control.Id == viewCodeMenuId);
            var beforeIndex = beforeControl == null ? 1 : beforeControl.Index;

            var items = GetFormDesignerContextMenuItems();
            BindParentMenuItem<FormDesignerContextParentMenu>(parent, beforeIndex, items);
        }

        private void ConfigureFormDesignerControlContextMenu()
        {
            const int viewCodeMenuId = 2558;
            var parent = Kernel.Get<VBE>().CommandBars[MsFormsControl].Controls;
            var beforeControl = parent.Cast<CommandBarControl>().FirstOrDefault(control => control.Id == viewCodeMenuId);
            var beforeIndex = beforeControl == null ? 1 : beforeControl.Index;

            var items = GetFormDesignerContextMenuItems();
            BindParentMenuItem<FormDesignerControlContextParentMenu>(parent, beforeIndex, items);
        }

        private void ConfigureProjectExplorerContextMenu()
        {
            const int projectPropertiesMenuId = 2578;
            var parent = Kernel.Get<VBE>().CommandBars[ProjectWindow].Controls;
            var beforeControl = parent.Cast<CommandBarControl>().FirstOrDefault(control => control.Id == projectPropertiesMenuId);
            var beforeIndex = beforeControl == null ? 1 : beforeControl.Index;

            var items = GetProjectWindowContextMenuItems();
            BindParentMenuItem<ProjectWindowContextParentMenu>(parent, beforeIndex, items);
        }

        private void BindParentMenuItem<TParentMenu>(CommandBarControls parent, int beforeIndex, IEnumerable<IMenuItem> items)
        {
            Bind<IParentMenuItem>().To(typeof(TParentMenu))
                .WhenInjectedInto<IAppMenu>()
                .InThreadScope()
                .WithConstructorArgument("items", items)
                .WithConstructorArgument("beforeIndex", beforeIndex)
                .WithPropertyValue("Parent", parent);
        }

        private static int FindRubberduckMenuInsertionIndex(CommandBarControls controls, int beforeId)
        {
            for (var i = 1; i <= controls.Count; i++)
            {
                if (controls[i].BuiltIn && controls[i].Id == beforeId)
                {
                    return i;
                }
            }

            return controls.Count;
        }

        private void BindCommandsToMenuItems()
        {
            var types = Assembly.GetExecutingAssembly().GetTypes()
                .Where(type => type.IsClass && type.Namespace != null && type.Namespace.StartsWith(typeof(CommandBase).Namespace ?? String.Empty))
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
                            .InSingletonScope();
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

        private IEnumerable<IMenuItem> GetRubberduckMenuItems()
        {
            return new[]
            {
                Kernel.Get<AboutCommandMenuItem>(),
                Kernel.Get<SettingsCommandMenuItem>(),
                Kernel.Get<InspectionResultsCommandMenuItem>(),
                GetUnitTestingParentMenu(),
                GetSmartIndenterParentMenu(),
                GetToolsParentMenu(),
                GetRefactoringsParentMenu(),
                GetNavigateParentMenu(),
            };
        }

        private IMenuItem GetUnitTestingParentMenu()
        {
            var items = new IMenuItem[]
            {
                Kernel.Get<RunAllTestsCommandMenuItem>(),
                Kernel.Get<TestExplorerCommandMenuItem>(),
                Kernel.Get<AddTestModuleCommandMenuItem>(),
                Kernel.Get<AddTestMethodCommandMenuItem>(),
                Kernel.Get<AddTestMethodExpectedErrorCommandMenuItem>(),
            };
            return new UnitTestingParentMenu(items);
        }

        private IMenuItem GetRefactoringsParentMenu()
        {
            var items = new IMenuItem[]
            {
                Kernel.Get<CodePaneRefactorRenameCommandMenuItem>(),
                Kernel.Get<RefactorExtractMethodCommandMenuItem>(),
                Kernel.Get<RefactorReorderParametersCommandMenuItem>(),
                Kernel.Get<RefactorRemoveParametersCommandMenuItem>(),
                Kernel.Get<RefactorIntroduceParameterCommandMenuItem>(),
                Kernel.Get<RefactorIntroduceFieldCommandMenuItem>(),
                Kernel.Get<RefactorEncapsulateFieldCommandMenuItem>(),
                Kernel.Get<RefactorMoveCloserToUsageCommandMenuItem>(),
                Kernel.Get<RefactorExtractInterfaceCommandMenuItem>(),
                Kernel.Get<RefactorImplementInterfaceCommandMenuItem>()
            };
            return new RefactoringsParentMenu(items);
        }

        private IMenuItem GetNavigateParentMenu()
        {
            var items = new IMenuItem[]
            {
                Kernel.Get<CodeExplorerCommandMenuItem>(),
                //Kernel.Get<RegexSearchReplaceCommandMenuItem>(),
                Kernel.Get<FindSymbolCommandMenuItem>(),
                Kernel.Get<FindAllReferencesCommandMenuItem>(),
                Kernel.Get<FindAllImplementationsCommandMenuItem>(),
            };
            return new NavigateParentMenu(items);
        }

        private IMenuItem GetSmartIndenterParentMenu()
        {
            var items = new IMenuItem[]
            {
                Kernel.Get<IndentCurrentProcedureCommandMenuItem>(),
                Kernel.Get<IndentCurrentModuleCommandMenuItem>(),
                Kernel.Get<NoIndentAnnotationCommandMenuItem>()
            };

            return new SmartIndenterParentMenu(items);
        }

        private IEnumerable<IMenuItem> GetCodePaneContextMenuItems()
        {
            return new[]
            {
                GetRefactoringsParentMenu(),
                GetSmartIndenterParentMenu(),
                Kernel.Get<FindSymbolCommandMenuItem>(),
                Kernel.Get<FindAllReferencesCommandMenuItem>(),
                Kernel.Get<FindAllImplementationsCommandMenuItem>(),
            };
        }

        private IMenuItem GetToolsParentMenu()
        {
            var items = new IMenuItem[]
            {
                Kernel.Get<ShowSourceControlPanelCommandMenuItem>(),
                Kernel.Get<RegexAssistantCommandMenuItem>(),
                Kernel.Get<ToDoExplorerCommandMenuItem>(),
            };

            return new ToolsParentMenu(items);
        }

        private IEnumerable<IMenuItem> GetFormDesignerContextMenuItems()
        {
            return new IMenuItem[]
            {
                Kernel.Get<FormDesignerRefactorRenameCommandMenuItem>(),
            };
        }

        private IEnumerable<IMenuItem> GetProjectWindowContextMenuItems()
        {
            return new IMenuItem[]
            {
                Kernel.Get<ProjectExplorerRefactorRenameCommandMenuItem>(),
                Kernel.Get<FindSymbolCommandMenuItem>(),
                Kernel.Get<FindAllReferencesCommandMenuItem>(),
                Kernel.Get<FindAllImplementationsCommandMenuItem>(),
            };
        }
    }
}
