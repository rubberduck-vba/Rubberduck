using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows.Input;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Ninject;
using Ninject.Extensions.Conventions;
using Ninject.Extensions.NamedScope;
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
using NLog;
using Rubberduck.Parsing.Preprocessing;
using System.Globalization;

namespace Rubberduck.Root
{
    public class RubberduckModule : NinjectModule
    {
        private readonly IKernel _kernel;
        private readonly VBE _vbe;
        private readonly AddIn _addin;

        private const int MenuBar = 1;
        private const int CodeWindow = 9;
        private const int ProjectWindow = 14;
        private const int MsForms = 17;
        private const int MsFormsControl = 18;

        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public RubberduckModule(IKernel kernel, VBE vbe, AddIn addin)
        {
            _kernel = kernel;
            _vbe = vbe;
            _addin = addin;
        }

        public override void Load()
        {
            _logger.Debug("in RubberduckModule.Load()");

            // bind VBE and AddIn dependencies to host-provided instances.
            _kernel.Bind<VBE>().ToConstant(_vbe);
            _kernel.Bind<AddIn>().ToConstant(_addin);
            _kernel.Bind<App>().ToSelf().InSingletonScope();
            _kernel.Bind<RubberduckParserState>().ToSelf().InSingletonScope();
            _kernel.Bind<GitProvider>().ToSelf().InSingletonScope();
            _kernel.Bind<NewUnitTestModuleCommand>().ToSelf().InSingletonScope();
            _kernel.Bind<NewTestMethodCommand>().ToSelf().InSingletonScope();
            _kernel.Bind<RubberduckCommandBar>().ToSelf().InSingletonScope();
            _kernel.Bind<TestExplorerModel>().ToSelf().InSingletonScope();

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
            
            Rebind<IIndenter>().To<Indenter>().InSingletonScope();
            Rebind<IIndenterSettings>().To<IndenterSettings>();
            Bind<Func<IIndenterSettings>>().ToMethod(t => () => _kernel.Get<IGeneralConfigService>().LoadConfiguration().UserSettings.IndenterSettings);

            //Bind<TestExplorerModel>().To<StandardModuleTestExplorerModel>().InSingletonScope();
            Rebind<IRubberduckParser>().To<RubberduckParser>().InSingletonScope();
            Bind<Func<IVBAPreprocessor>>().ToMethod(p => () => new VBAPreprocessor(double.Parse(_vbe.Version, CultureInfo.InvariantCulture)));

            _kernel.Rebind<ISearchResultsWindowViewModel>().To<SearchResultsWindowViewModel>().InSingletonScope();
            _kernel.Bind<SearchResultPresenterInstanceManager>().ToSelf().InSingletonScope();

            Bind<IPresenter>().To<TestExplorerDockablePresenter>()
                .WhenInjectedInto<TestExplorerCommand>()
                .InSingletonScope()
                .WithConstructorArgument<IDockableUserControl>(new TestExplorerWindow { ViewModel = _kernel.Get<TestExplorerViewModel>() });

            Bind<IPresenter>().To<CodeInspectionsDockablePresenter>()
                .WhenInjectedInto<InspectionResultsCommand>()
                .InSingletonScope()
                .WithConstructorArgument<IDockableUserControl>(new CodeInspectionsWindow { ViewModel = _kernel.Get<InspectionResultsViewModel>() });

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
                .WithConstructorArgument(new SourceControlPanel { ViewModel = _kernel.Get<SourceControlViewViewModel>() });
            
            BindCommandsToCodeExplorer();
            Bind<IPresenter>().To<CodeExplorerDockablePresenter>()
                .WhenInjectedInto<CodeExplorerCommand>()
                .InSingletonScope()
                .WithConstructorArgument<IDockableUserControl>(new CodeExplorerWindow { ViewModel = _kernel.Get<CodeExplorerViewModel>() });

            Bind<IPresenter>().To<ToDoExplorerDockablePresenter>()
                .WhenInjectedInto<ToDoExplorerCommand>()
                .InSingletonScope()
                .WithConstructorArgument<IDockableUserControl>(new ToDoExplorerWindow { ViewModel = _kernel.Get<ToDoExplorerViewModel>() });

            ConfigureRubberduckMenu();
            ConfigureCodePaneContextMenu();
            ConfigureFormDesignerContextMenu();
            ConfigureFormDesignerControlContextMenu();
            ConfigureProjectExplorerContextMenu();

            BindWindowsHooks();
            _logger.Debug("completed RubberduckModule.Load()");
        }

        private void BindWindowsHooks()
        {
            _kernel.Rebind<IAttachable>().To<TimerHook>()
                .InSingletonScope()
                .WithConstructorArgument("mainWindowHandle", (IntPtr)_vbe.MainWindow.HWnd);

            _kernel.Rebind<IRubberduckHooks>().To<RubberduckHooks>()
                .InSingletonScope()
                .WithConstructorArgument("mainWindowHandle", (IntPtr)_vbe.MainWindow.HWnd);
        }

        private void ApplyDefaultInterfacesConvention(IEnumerable<Assembly> assemblies)
        {
            _kernel.Bind(t => t.From(assemblies)
                .SelectAllClasses()
                // inspections & factories have their own binding rules
                .Where(type => !type.Name.EndsWith("Factory") && !type.Name.EndsWith("ConfigProvider") && !type.GetInterfaces().Contains(typeof(IInspection)))
                .BindDefaultInterface()
                .Configure(binding => binding.InCallScope())); // TransientScope wouldn't dispose disposables
        }

        // note: settings namespace classes are injected in singleton scope
        private void ApplyConfigurationConvention(IEnumerable<Assembly> assemblies)
        {
            _kernel.Bind(t => t.From(assemblies)
                .SelectAllClasses()
                .InNamespaceOf<Configuration>()
                .BindAllInterfaces()
                .Configure(binding => binding.InSingletonScope()));

            _kernel.Bind<IPersistanceService<CodeInspectionSettings>>().To<XmlPersistanceService<CodeInspectionSettings>>().InSingletonScope();
            _kernel.Bind<IPersistanceService<GeneralSettings>>().To<XmlPersistanceService<GeneralSettings>>().InSingletonScope();
            _kernel.Bind<IPersistanceService<HotkeySettings>>().To<XmlPersistanceService<HotkeySettings>>().InSingletonScope();
            _kernel.Bind<IPersistanceService<ToDoListSettings>>().To<XmlPersistanceService<ToDoListSettings>>().InSingletonScope();
            _kernel.Bind<IPersistanceService<UnitTestSettings>>().To<XmlPersistanceService<UnitTestSettings>>().InSingletonScope();
            _kernel.Bind<IPersistanceService<IndenterSettings>>().To<XmlPersistanceService<IndenterSettings>>().InSingletonScope();
            _kernel.Bind<IFilePersistanceService<SourceControlSettings>>().To<XmlPersistanceService<SourceControlSettings>>().InSingletonScope();

            _kernel.Bind<IIndenterConfigProvider>().To<IndenterConfigProvider>().InSingletonScope();
            _kernel.Bind<ISourceControlConfigProvider>().To<SourceControlConfigProvider>().InSingletonScope();

            _kernel.Bind<ICodeInspectionSettings>().To<CodeInspectionSettings>();
            _kernel.Bind<IGeneralSettings>().To<GeneralSettings>();
            _kernel.Bind<IHotkeySettings>().To<HotkeySettings>();
            _kernel.Bind<IToDoListSettings>().To<ToDoListSettings>();
            _kernel.Bind<IUnitTestSettings>().To<UnitTestSettings>();
            _kernel.Bind<IIndenterSettings>().To<IndenterSettings>();
            _kernel.Bind<ISourceControlSettings>().To<SourceControlSettings>();        
        }

        // note convention: abstract factory interface names end with "Factory".
        private void ApplyAbstractFactoryConvention(IEnumerable<Assembly> assemblies)
        {
            _kernel.Bind(t => t.From(assemblies)
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
                _kernel.Bind<IInspection>().To(inspection).InSingletonScope();
            }
        }

        private void ConfigureRubberduckMenu()
        {
            const int windowMenuId = 30009;
            var parent = _kernel.Get<VBE>().CommandBars[MenuBar].Controls;
            var beforeIndex = FindRubberduckMenuInsertionIndex(parent, windowMenuId);

            var items = GetRubberduckMenuItems();
            BindParentMenuItem<RubberduckParentMenu>(parent, beforeIndex, items);
        }

        private void ConfigureCodePaneContextMenu()
        {
            const int listMembersMenuId = 2529;
            var parent = _kernel.Get<VBE>().CommandBars[CodeWindow].Controls;
            var beforeControl = parent.Cast<CommandBarControl>().FirstOrDefault(control => control.Id == listMembersMenuId);
            var beforeIndex = beforeControl == null ? 1 : beforeControl.Index;

            var items = GetCodePaneContextMenuItems();
            BindParentMenuItem<CodePaneContextParentMenu>(parent, beforeIndex, items);
        }

        private void ConfigureFormDesignerContextMenu()
        {
            const int viewCodeMenuId = 2558;
            var parent = _kernel.Get<VBE>().CommandBars[MsForms].Controls;
            var beforeControl = parent.Cast<CommandBarControl>().FirstOrDefault(control => control.Id == viewCodeMenuId);
            var beforeIndex = beforeControl == null ? 1 : beforeControl.Index;

            var items = GetFormDesignerContextMenuItems();
            BindParentMenuItem<FormDesignerContextParentMenu>(parent, beforeIndex, items);
        }

        private void ConfigureFormDesignerControlContextMenu()
        {
            const int viewCodeMenuId = 2558;
            var parent = _kernel.Get<VBE>().CommandBars[MsFormsControl].Controls;
            var beforeControl = parent.Cast<CommandBarControl>().FirstOrDefault(control => control.Id == viewCodeMenuId);
            var beforeIndex = beforeControl == null ? 1 : beforeControl.Index;

            var items = GetFormDesignerContextMenuItems();
            BindParentMenuItem<FormDesignerControlContextParentMenu>(parent, beforeIndex, items);
        }

        private void ConfigureProjectExplorerContextMenu()
        {
            const int projectPropertiesMenuId = 2578;
            var parent = _kernel.Get<VBE>().CommandBars[ProjectWindow].Controls;
            var beforeControl = parent.Cast<CommandBarControl>().FirstOrDefault(control => control.Id == projectPropertiesMenuId);
            var beforeIndex = beforeControl == null ? 1 : beforeControl.Index;

            var items = GetProjectWindowContextMenuItems();
            BindParentMenuItem<ProjectWindowContextParentMenu>(parent, beforeIndex, items);
        }

        private void BindParentMenuItem<TParentMenu>(CommandBarControls parent, int beforeIndex, IEnumerable<IMenuItem> items)
        {
            _kernel.Bind<IParentMenuItem>().To(typeof(TParentMenu))
                .WhenInjectedInto<IAppMenu>()
                .InCallScope()
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

            // note: ICommand naming convention: [Foo]Command
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
                        var binding = _kernel.Bind<ICommand>().To(command);
                        var whenCommandMenuItemCondition =
                            binding.WhenInjectedInto(item).BindingConfiguration.Condition;
                        var whenHooksCondition =
                            binding.WhenInjectedInto<RubberduckHooks>().BindingConfiguration.Condition;

                        binding.When(request => whenCommandMenuItemCondition(request) || whenHooksCondition(request))
                               .InSingletonScope();

                        //_kernel.Bind<ICommand>().To(command).WhenInjectedExactlyInto(item);
                        //_kernel.Bind<ICommand>().To(command);
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
                               type.Namespace == "Rubberduck.UI.CodeExplorer.Commands" &&
                               type.Name.StartsWith("CodeExplorer_") && 
                               type.Name.EndsWith("Command"));

            foreach (var command in commands)
            {
                _kernel.Bind<ICommand>().To(command).InSingletonScope();
            }
        }

        private IEnumerable<IMenuItem> GetRubberduckMenuItems()
        {
            return new[]
            {
                _kernel.Get<AboutCommandMenuItem>(),
                _kernel.Get<SettingsCommandMenuItem>(),
                _kernel.Get<InspectionResultsCommandMenuItem>(),
                _kernel.Get<ShowSourceControlPanelCommandMenuItem>(),
                GetUnitTestingParentMenu(),
                GetSmartIndenterParentMenu(),
                GetRefactoringsParentMenu(),
                GetNavigateParentMenu(),
            };
        }

        private IMenuItem GetUnitTestingParentMenu()
        {
            var items = new IMenuItem[]
            {
                _kernel.Get<RunAllTestsCommandMenuItem>(),
                _kernel.Get<TestExplorerCommandMenuItem>(),
                _kernel.Get<AddTestModuleCommandMenuItem>(),
                _kernel.Get<AddTestMethodCommandMenuItem>(),
                _kernel.Get<AddTestMethodExpectedErrorCommandMenuItem>(),
            };
            return new UnitTestingParentMenu(items);
        }

        private IMenuItem GetRefactoringsParentMenu()
        {
            var items = new IMenuItem[]
            {
                _kernel.Get<CodePaneRefactorRenameCommandMenuItem>(),
                _kernel.Get<RefactorExtractMethodCommandMenuItem>(),
                _kernel.Get<RefactorReorderParametersCommandMenuItem>(),
                _kernel.Get<RefactorRemoveParametersCommandMenuItem>(),
                _kernel.Get<RefactorIntroduceParameterCommandMenuItem>(),
                _kernel.Get<RefactorIntroduceFieldCommandMenuItem>(),
                _kernel.Get<RefactorEncapsulateFieldCommandMenuItem>(),
                _kernel.Get<RefactorMoveCloserToUsageCommandMenuItem>(),
                _kernel.Get<RefactorExtractInterfaceCommandMenuItem>(),
                _kernel.Get<RefactorImplementInterfaceCommandMenuItem>()
            };
            return new RefactoringsParentMenu(items);
        }

        private IMenuItem GetNavigateParentMenu()
        {
            var items = new IMenuItem[]
            {
                _kernel.Get<CodeExplorerCommandMenuItem>(),
                _kernel.Get<ToDoExplorerCommandMenuItem>(),
                //_kernel.Get<RegexSearchReplaceCommandMenuItem>(),
                _kernel.Get<FindSymbolCommandMenuItem>(),
                _kernel.Get<FindAllReferencesCommandMenuItem>(),
                _kernel.Get<FindAllImplementationsCommandMenuItem>(),
            };
            return new NavigateParentMenu(items);
        }

        private IMenuItem GetSmartIndenterParentMenu()
        {
            var items = new IMenuItem[]
            {
                _kernel.Get<IndentCurrentProcedureCommandMenuItem>(),
                _kernel.Get<IndentCurrentModuleCommandMenuItem>(),
                _kernel.Get<NoIndentAnnotationCommandMenuItem>()
            };

            return new SmartIndenterParentMenu(items);
        }

        private IEnumerable<IMenuItem> GetCodePaneContextMenuItems()
        {
            return new[]
            {
                GetRefactoringsParentMenu(),
                GetSmartIndenterParentMenu(),
                //_kernel.Get<RegexSearchReplaceCommandMenuItem>(),
                _kernel.Get<FindSymbolCommandMenuItem>(),
                _kernel.Get<FindAllReferencesCommandMenuItem>(),
                _kernel.Get<FindAllImplementationsCommandMenuItem>(),
            };
        }

        private IEnumerable<IMenuItem> GetFormDesignerContextMenuItems()
        {
            return new IMenuItem[]
            {
                _kernel.Get<FormDesignerRefactorRenameCommandMenuItem>(),
            };
        }

        private IEnumerable<IMenuItem> GetProjectWindowContextMenuItems()
        {
            return new IMenuItem[]
            {
                _kernel.Get<ProjectExplorerRefactorRenameCommandMenuItem>(),
                _kernel.Get<FindSymbolCommandMenuItem>(),
                _kernel.Get<FindAllReferencesCommandMenuItem>(),
                _kernel.Get<FindAllImplementationsCommandMenuItem>(),
            };
        }
    }
}
