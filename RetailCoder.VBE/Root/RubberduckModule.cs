using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Ninject;
using Ninject.Extensions.Conventions;
using Ninject.Modules;
using Rubberduck.Common;
using Rubberduck.Inspections;
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
using Rubberduck.Parsing.Preprocessing;
using System.Globalization;
using Ninject.Extensions.Interception.Infrastructure.Language;
using Ninject.Extensions.NamedScope;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

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
            Bind<Sinks>().ToSelf().InSingletonScope();
            Bind<App>().ToSelf().InSingletonScope();
            Bind<RubberduckParserState>().ToSelf().InSingletonScope();
            Bind<GitProvider>().ToSelf().InSingletonScope();
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

            ApplyDefaultInterfacesConvention(assemblies);
            ApplyConfigurationConvention(assemblies);
            ApplyAbstractFactoryConvention(assemblies);

            BindCommandsToMenuItems();

            Rebind<ISinks>().To<Sinks>().InSingletonScope();
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
                .InSingletonScope();

            Bind<IPresenter>().To<CodeInspectionsDockablePresenter>()
                .WhenInjectedInto<InspectionResultsCommand>()
                .InSingletonScope();

            Bind<IControlView>().To<ChangesView>().InCallScope();
            Bind<IControlView>().To<BranchesView>().InCallScope();
            Bind<IControlView>().To<UnsyncedCommitsView>().InCallScope();
            Bind<IControlView>().To<SettingsView>().InCallScope();

            Bind<IControlViewModel>().To<ChangesViewViewModel>()
                .WhenInjectedInto<ChangesView>().InCallScope();
            Bind<IControlViewModel>().To<BranchesViewViewModel>()
                .WhenInjectedInto<BranchesView>().InCallScope();
            Bind<IControlViewModel>().To<UnsyncedCommitsViewViewModel>()
                .WhenInjectedInto<UnsyncedCommitsView>().InCallScope();
            Bind<IControlViewModel>().To<SettingsViewViewModel>()
                .WhenInjectedInto<SettingsView>().InCallScope();

            Bind<ISourceControlProviderFactory>().To<SourceControlProviderFactory>()
                .WhenInjectedInto<SourceControlViewViewModel>();

            Bind<SourceControlDockablePresenter>().ToSelf().InSingletonScope();
            
            BindCommandsToCodeExplorer();
            Bind<IPresenter>().To<CodeExplorerDockablePresenter>()
                .WhenInjectedInto<CodeExplorerCommand>()
                .InSingletonScope();

            Bind<IPresenter>().To<ToDoExplorerDockablePresenter>()
                .WhenInjectedInto<ToDoExplorerCommand>()
                .InSingletonScope();

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
                .Where(type => type.Namespace != null
                            && !type.Namespace.StartsWith("Rubberduck.VBEditor.SafeComWrappers")
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

            Bind<IPersistanceService<CodeInspectionSettings>>().To<XmlPersistanceService<CodeInspectionSettings>>().InCallScope();
            Bind<IPersistanceService<GeneralSettings>>().To<XmlPersistanceService<GeneralSettings>>().InCallScope();
            Bind<IPersistanceService<HotkeySettings>>().To<XmlPersistanceService<HotkeySettings>>().InCallScope();
            Bind<IPersistanceService<ToDoListSettings>>().To<XmlPersistanceService<ToDoListSettings>>().InCallScope();
            Bind<IPersistanceService<UnitTestSettings>>().To<XmlPersistanceService<UnitTestSettings>>().InCallScope();
            Bind<IPersistanceService<IndenterSettings>>().To<XmlPersistanceService<IndenterSettings>>().InCallScope();
            Bind<IFilePersistanceService<SourceControlSettings>>().To<XmlPersistanceService<SourceControlSettings>>().InCallScope();

            Bind<IConfigProvider<IndenterSettings>>().To<IndenterConfigProvider>().InCallScope();
            Bind<IConfigProvider<SourceControlSettings>>().To<SourceControlConfigProvider>().InCallScope();

            Bind<ICodeInspectionSettings>().To<CodeInspectionSettings>().InCallScope();
            Bind<IGeneralSettings>().To<GeneralSettings>().InCallScope();
            Bind<IHotkeySettings>().To<HotkeySettings>().InCallScope();
            Bind<IToDoListSettings>().To<ToDoListSettings>().InCallScope();
            Bind<IUnitTestSettings>().To<UnitTestSettings>().InCallScope();
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
                        .InCallScope()
                        .Named(inspection.FullName);

                    binding.Intercept().With<TimedCallLoggerInterceptor>();
                    binding.Intercept().With<EnumerableCounterInterceptor<InspectionResultBase>>();

                    var localInspection = inspection;
                    Bind<IInspection>().ToMethod(
                        c => c.Kernel.Get<IParseTreeInspection>(localInspection.FullName));
                }
                else
                {
                    var binding = Bind<IInspection>().To(inspection).InCallScope();
                    binding.Intercept().With<TimedCallLoggerInterceptor>();
                    binding.Intercept().With<EnumerableCounterInterceptor<InspectionResultBase>>();
                }
            }
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
