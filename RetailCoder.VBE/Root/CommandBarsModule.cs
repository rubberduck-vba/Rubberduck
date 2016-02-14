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
using Ninject.Parameters;
using Rubberduck.Navigation;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.Root
{
    public class CommandBarsModule : NinjectModule
    {
        private const int MenuBar = 1;
        private const int CodeWindow = 9;
        private const int ProjectWindow = 14;
        private const int MsForms = 17;
        private const int MsFormsControl = 18;

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
            _kernel.Bind<IDeclarationNavigator>().To<NavigateAllImplementations>().WhenTargetHas<FindImplementationsAttribute>();
            _kernel.Bind<IDeclarationNavigator>().To<NavigateAllReferences>().WhenTargetHas<FindReferencesAttribute>();

            var types = Assembly.GetExecutingAssembly().GetTypes()
                .Where(type => type.IsClass && type.Namespace != null && type.Namespace.StartsWith(typeof(CommandBase).Namespace ?? string.Empty))
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
                        _kernel.Bind<ICommand>().To(command).WhenInjectedExactlyInto(item).InSingletonScope();
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
                _kernel.Get<RegexSearchReplaceCommandMenuItem>(),
                _kernel.Get<FindSymbolCommandMenuItem>(),
                _kernel.Get<FindAllReferencesCommandMenuItem>(),
                _kernel.Get<FindAllImplementationsCommandMenuItem>(),
                _kernel.Get<RegexSearchReplaceCommandMenuItem>(),
            };
            return new NavigateParentMenu(items);
        }

        private IMenuItem GetSmartIndenterParentMenu()
        {
            var items = new IMenuItem[]
            {
                _kernel.Get<IndentCurrentProcedureCommandMenuItem>(),
                _kernel.Get<IndentCurrentModuleCommandMenuItem>()
            };

            return new SmartIndenterParentMenu(items);
        }

        private IEnumerable<IMenuItem> GetCodePaneContextMenuItems()
        {
            return new IMenuItem[]
            {
                GetRefactoringsParentMenu(),
                GetSmartIndenterParentMenu(),
                _kernel.Get<RegexSearchReplaceCommandMenuItem>(),
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
