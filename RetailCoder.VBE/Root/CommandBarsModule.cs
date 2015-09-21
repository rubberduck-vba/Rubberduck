using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Input;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Ninject;
using Ninject.Modules;
using Rubberduck.Navigation;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

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
                .InSingletonScope()
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
            //_kernel.Bind<ICommand>().To<NavigateCommand>().InSingletonScope();
            _kernel.Bind<IDeclarationNavigator>().To<NavigateAllImplementations>().WhenTargetHas<FindImplementationsAttribute>().InSingletonScope();
            _kernel.Bind<IDeclarationNavigator>().To<NavigateAllReferences>().WhenTargetHas<FindReferencesAttribute>().InSingletonScope();

            var types = Assembly.GetExecutingAssembly().GetTypes()
                .Where(type => type.Namespace != null && type.Namespace.StartsWith(typeof(CommandBase).Namespace ?? string.Empty))
                .ToList();

            // note: ICommand naming convention: [Foo]Command
            var commands = types.Where(type => type.IsClass && type.BaseType == typeof(CommandBase) && type.Name.EndsWith("Command"));
            foreach (var command in commands)
            {
                var commandName = command.Name.Substring(0, command.Name.Length - "Command".Length);
                try
                {
                    // note: ICommandMenuItem naming convention for [Foo]Command: [Foo]CommandMenuItem
                    var item = types.SingleOrDefault(type => type.Name == commandName + "CommandMenuItem");
                    if (item != null)
                    {
                        _kernel.Bind(item).ToSelf().InSingletonScope();
                        _kernel.Bind<ICommand>().To(command).WhenInjectedInto(item).InSingletonScope();
                    }
                }
                catch (InvalidOperationException exception)
                {
                    // rename one of the classes, "FooCommand" is expected to match exactly 1 "FooBarXyzCommandMenuItem"
                }
            }
        }

        private IEnumerable<IMenuItem> GetRubberduckMenuItems()
        {
            return new IMenuItem[]
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
                _kernel.Get<RefactorRenameCommandMenuItem>(), 
                _kernel.Get<RefactorExtractMethodCommandMenuItem>(), 
                _kernel.Get<RefactorReorderParametersCommandMenuItem>(), 
                _kernel.Get<RefactorRemoveParametersCommandMenuItem>(), 
            };

            return new RefactoringsParentMenu(items);
        }

        private IMenuItem GetNavigateParentMenu()
        {
            var items = new IMenuItem[]
            {
                _kernel.Get<CodeExplorerCommandMenuItem>(),
                _kernel.Get<ToDoExplorerCommandMenuItem>(),
                _kernel.Get<FindSymbolCommandMenuItem>(),
                _kernel.Get<FindAllReferencesCommandMenuItem>(),
                _kernel.Get<FindAllImplementationsCommandMenuItem>(),
            };
            return new NavigateParentMenu(items);
        }

        private IEnumerable<IMenuItem> GetCodePaneContextMenuItems()
        {
            return new IMenuItem[]
            {
                GetRefactoringsParentMenu(),
                _kernel.Get<FindSymbolCommandMenuItem>(),
                _kernel.Get<FindAllReferencesCommandMenuItem>(),
                _kernel.Get<FindAllImplementationsCommandMenuItem>(),
            };
        }

        private IEnumerable<IMenuItem> GetFormDesignerContextMenuItems()
        {
            return new IMenuItem[]
            {
                _kernel.Get<RefactorRenameCommandMenuItem>(), 
            };
        }

        private IEnumerable<IMenuItem> GetProjectWindowContextMenuItems()
        {
            return new IMenuItem[]
            {
                _kernel.Get<RefactorRenameCommandMenuItem>(), 
                _kernel.Get<FindSymbolCommandMenuItem>(),
                _kernel.Get<FindAllReferencesCommandMenuItem>(),
                _kernel.Get<FindAllImplementationsCommandMenuItem>(),
            };
        }
    }
}
