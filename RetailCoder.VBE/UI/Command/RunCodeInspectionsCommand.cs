using System.Runtime.InteropServices;
using System.Windows.Input;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that runs all active code inspections for the active VBAProject.
    /// </summary>
    [ComVisible(false)]
    public class RunCodeInspectionsCommand : CommandBase
    {
        public override void Execute(object parameter)
        {
            throw new System.NotImplementedException();
        }
    }

    public class RunCodeInspectionsCommandMenuItem : CommandMenuItemBase
    {
        public RunCodeInspectionsCommandMenuItem(ICommand command) 
            : base(command)
        {
        }

        public override string Key { get { return "RubberduckMenu_CodeInspections"; } }
        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.CodeInspections; } }
    }
}