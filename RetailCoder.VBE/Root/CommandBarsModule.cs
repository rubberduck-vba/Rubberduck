using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Input;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Ninject;
using Ninject.Extensions.NamedScope;
using Ninject.Modules;
using Rubberduck.Navigation;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.Root
{
    public class CommandBarsModule : NinjectModule
    {
        private const int MENU_BAR = 1;
        private const int CODE_WINDOW = 9;
        private const int PROJECT_WINDOW = 14;
        private const int MS_FORMS = 17;
        private const int MS_FORMS_CONTROL = 18;

        private readonly IKernel _kernel;

        public CommandBarsModule(IKernel kernel)
        {
            _kernel = kernel;
        }

        public override void Load()
        {
            BindCommandsToMenuItems();

            ConfigureRubberduckMenu();
            ConfigureCodePaneContextMenu();
            ConfigureFormDesignerContextMenu();
            ConfigureFormDesignerControlContextMenu();
            ConfigureProjectExplorerContextMenu();
        }

        private void ConfigureRubberduckMenu()
        {
            const int windowMenuId = 30009;
            var parent = _kernel.Get<VBE>().CommandBars[MENU_BAR].Controls;
            var beforeIndex = FindRubberduckMenuInsertionIndex(parent, windowMenuId);

            var items = GetRubberduckMenuItems();
            BindParentMenuItem<RubberduckParentMenu>(parent, beforeIndex, items);
        }

        private void ConfigureCodePaneContextMenu()
        {
            const int listMembersMenuId = 2529;
            var parent = _kernel.Get<VBE>().CommandBars[CODE_WINDOW].Controls;
            var beforeIndex = parent.Cast<CommandBarControl>().First(control => control.Id == listMembersMenuId).Index;

            var items = GetCodePaneContextMenuItems();
            BindParentMenuItem<RubberduckParentMenu>(parent, beforeIndex, items);
        }

        private void ConfigureFormDesignerContextMenu()
        {
            const int viewCodeMenuId = 2558;
            var parent = _kernel.Get<VBE>().CommandBars[MS_FORMS].Controls;
            var beforeIndex = parent.Cast<CommandBarControl>().First(control => control.Id == viewCodeMenuId).Index;

            var items = GetFormDesignerContextMenuItems();
            BindParentMenuItem<FormDesignerContextParentMenu>(parent, beforeIndex, items);
        }

        private void ConfigureFormDesignerControlContextMenu()
        {
            const int viewCodeMenuId = 2558;
            var parent = _kernel.Get<VBE>().CommandBars[MS_FORMS_CONTROL].Controls;
            var beforeIndex = parent.Cast<CommandBarControl>().First(control => control.Id == viewCodeMenuId).Index;

            var items = GetFormDesignerContextMenuItems();
            BindParentMenuItem<FormDesignerControlContextParentMenu>(parent, beforeIndex, items);
        }

        private void ConfigureProjectExplorerContextMenu()
        {
            const int projectPropertiesMenuId = 2578;
            var parent = _kernel.Get<VBE>().CommandBars[PROJECT_WINDOW].Controls;
            var beforeIndex = parent.Cast<CommandBarControl>().First(control => control.Id == projectPropertiesMenuId).Index;

            var items = GetProjectWindowContextMenuItems();
            BindParentMenuItem<ProjectWindowContextParentMenu>(parent, beforeIndex, items);
        }

        private void BindParentMenuItem<TParentMenu>(CommandBarControls parent, int beforeIndex, IEnumerable<IMenuItem> items)
        {
            _kernel.Bind<IParentMenuItem>().To(typeof(TParentMenu))
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
            _kernel.Bind<IDeclarationNavigator>().To<NavigateAllImplementations>().WhenTargetHas<FindImplementationsAttribute>();
            _kernel.Bind<IDeclarationNavigator>().To<NavigateAllReferences>().WhenTargetHas<FindReferencesAttribute>();

            var types = Assembly.GetExecutingAssembly().GetTypes()
                .Where(type => type.Namespace != null && type.Namespace.StartsWith(typeof(CommandBase).Namespace ?? string.Empty))
                .ToList();

            // note: ICommand naming convention: [Foo]Command
            var baseCommandTypes = new[] {typeof (CommandBase), typeof (RefactorCommandBase)};
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
                        _kernel.Bind(item).ToSelf().InCallScope();
                        _kernel.Bind<ICommand>().To(command).WhenInjectedInto(item).InCallScope();
                    }
                }
                catch (InvalidOperationException exception)
                {
                    // rename one of the classes, "FooCommand" is expected to match exactly 1 "FooBarXyzCommandMenuItem"
                }
            }
        }

        private IEnumerable<IMenuItem> _rubberduckMenuItems;
        private IEnumerable<IMenuItem> GetRubberduckMenuItems()
        {
            if (_rubberduckMenuItems == null)
            {
                _rubberduckMenuItems = new[]
                {
                    _kernel.Get<AboutCommandMenuItem>(),
                    _kernel.Get<OptionsCommandMenuItem>(),
                    _kernel.Get<RunCodeInspectionsCommandMenuItem>(),
                    _kernel.Get<ShowSourceControlPanelCommandMenuItem>(),
                    GetUnitTestingParentMenu(),
                    GetRefactoringsParentMenu(),
                    GetNavigateParentMenu(),
                };
            }
            return _rubberduckMenuItems;
        }

        private UnitTestingParentMenu _unitTestingParentMenu;
        private IMenuItem GetUnitTestingParentMenu()
        {
            if (_unitTestingParentMenu == null)
            {
                var items = new IMenuItem[]
                {
                    _kernel.Get<RunAllTestsCommandMenuItem>(),
                    _kernel.Get<TestExplorerCommandMenuItem>(),
                    _kernel.Get<AddTestModuleCommandMenuItem>(),
                    _kernel.Get<AddTestMethodCommandMenuItem>(),
                    _kernel.Get<AddTestMethodExpectedErrorCommandMenuItem>(),
                };
                _unitTestingParentMenu = new UnitTestingParentMenu(items.ToList());
            }
            return _unitTestingParentMenu;
        }

        private RefactoringsParentMenu _refactoringsParentMenu;
        private IMenuItem GetRefactoringsParentMenu()
        {
            if (_refactoringsParentMenu == null)
            {
                var items = new IMenuItem[]
                {
                    _kernel.Get<RefactorRenameCommandMenuItem>(),
                    _kernel.Get<RefactorExtractMethodCommandMenuItem>(),
                    _kernel.Get<RefactorReorderParametersCommandMenuItem>(),
                    _kernel.Get<RefactorRemoveParametersCommandMenuItem>(),
                };
                _refactoringsParentMenu = new RefactoringsParentMenu(items.ToList());
            }
            return _refactoringsParentMenu;
        }

        private NavigateParentMenu _navigateParentMenu;
        private IMenuItem GetNavigateParentMenu()
        {
            if (_navigateParentMenu == null)
            {
                var items = new IMenuItem[]
                {
                    _kernel.Get<CodeExplorerCommandMenuItem>(),
                    _kernel.Get<ToDoExplorerCommandMenuItem>(),
                    _kernel.Get<RegexSearchReplaceCommandMenuItem>(),
                    _kernel.Get<FindSymbolCommandMenuItem>(),
                    _kernel.Get<FindAllReferencesCommandMenuItem>(),
                    _kernel.Get<FindAllImplementationsCommandMenuItem>(),
                };
                _navigateParentMenu = new NavigateParentMenu(items.ToList());
            }
            return _navigateParentMenu;
        }

        private IEnumerable<IMenuItem> _codePaneContextMenuItems; 
        private IEnumerable<IMenuItem> GetCodePaneContextMenuItems()
        {
            if (_codePaneContextMenuItems == null)
            {
                _codePaneContextMenuItems = new[]
                {
                    GetRefactoringsParentMenu(),
                    _kernel.Get<RegexSearchReplaceCommandMenuItem>(),
                    _kernel.Get<FindSymbolCommandMenuItem>(),
                    _kernel.Get<FindAllReferencesCommandMenuItem>(),
                    _kernel.Get<FindAllImplementationsCommandMenuItem>(),
                };
            }
            return _codePaneContextMenuItems.ToList();
        }

        private IEnumerable<IMenuItem> _formDesignerContextMenuItems; 
        private IEnumerable<IMenuItem> GetFormDesignerContextMenuItems()
        {
            if (_formDesignerContextMenuItems == null)
            {
                _formDesignerContextMenuItems = new IMenuItem[]
                {
                    _kernel.Get<RefactorRenameCommandMenuItem>(),
                };
            }
            return _formDesignerContextMenuItems.ToList();
        }

        private IEnumerable<IMenuItem> _projectWindowContextMenuItems; 
        private IEnumerable<IMenuItem> GetProjectWindowContextMenuItems()
        {
            if (_projectWindowContextMenuItems == null)
            {
                _projectWindowContextMenuItems = new IMenuItem[]
                {
                    _kernel.Get<RefactorRenameCommandMenuItem>(),
                    _kernel.Get<FindSymbolCommandMenuItem>(),
                    _kernel.Get<FindAllReferencesCommandMenuItem>(),
                    _kernel.Get<FindAllImplementationsCommandMenuItem>(),
                };
            }
            return _projectWindowContextMenuItems.ToList();
        }
    }
}
