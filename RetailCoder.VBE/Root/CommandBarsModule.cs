using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Ninject;
using Ninject.Activation;
using Ninject.Extensions.NamedScope;
using Ninject.Modules;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;

namespace Rubberduck.Root
{
    public class CommandBarsModule : NinjectModule
    {
        private readonly IKernel _kernel;

        public CommandBarsModule(IKernel kernel)
        {
            _kernel = kernel;
        }

        public override void Load()
        {
            ConfigureRubberduckMenu();
        }

        private void ConfigureRubberduckMenu()
        {
            BindCommandsToMenuItems();

            const int windowMenuId = 30009;
            var menuBarControls = _kernel.Get<VBE>().CommandBars[1].Controls;
            var beforeIndex = FindMenuInsertionIndex(menuBarControls, windowMenuId);
            var items = GetRubberduckMenuItems();

            _kernel.Bind<RubberduckParentMenu>().ToSelf().InSingletonScope()
                .WithConstructorArgument("items", items)
                .WithConstructorArgument("beforeIndex", beforeIndex)
                .WithPropertyValue("Parent", menuBarControls);
        }

        private static int FindMenuInsertionIndex(CommandBarControls controls, int beforeId)
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
                .Where(type => type.Namespace == typeof(ICommand).Namespace)
                .ToList();

            var commands = types.Where(type => type.IsClass && type.GetInterfaces().Contains(typeof(ICommand)) && type.Name.EndsWith("Command"));
            foreach (var command in commands)
            {
                var commandName = command.Name.Substring(0, command.Name.Length - "Command".Length);
                var item = types.SingleOrDefault(type => type.Name.StartsWith(commandName) && type.Name.EndsWith("CommandMenuItem"));
                if (item != null)
                {
                    _kernel.Bind(item).ToSelf().InSingletonScope();
                    _kernel.Bind<ICommand>().To(command).WhenInjectedInto(item).InSingletonScope();
                }
            }
        }

        private IEnumerable<IMenuItem> GetRubberduckMenuItems()
        {
            return new[]
            {
                _kernel.Get<AboutCommandMenuItem>(),
                _kernel.Get<CodeExplorerCommandMenuItem>(), 
                _kernel.Get<OptionsCommandMenuItem>(), 
                _kernel.Get<RunCodeInspectionsCommandMenuItem>(),
                _kernel.Get<ToDoExplorerCommandMenuItem>(), 
                _kernel.Get<ShowSourceControlPanelCommandMenuItem>(), 
                GetUnitTestingParentMenu(),
                GetRefactoringsParentMenu(),
            };
        }

        private IMenuItem GetUnitTestingParentMenu()
        {
            var items = new IMenuItem[]
            {
                _kernel.Get<RunAllTestsUnitTestingCommandMenuItem>(), 
                _kernel.Get<TestExplorerUnitTestingCommandMenuItem>(), 
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
    }
}
