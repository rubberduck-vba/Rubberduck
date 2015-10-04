using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Rubberduck.Inspections;
using Rubberduck.UI.CodeInspections;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that runs all active code inspections for the active VBAProject.
    /// </summary>
    [ComVisible(false)]
    public class RunCodeInspectionsCommand : CommandBase
    {
        private readonly IPresenter _presenter;

        public RunCodeInspectionsCommand(IPresenter presenter)
        {
            _presenter = presenter;
        }

        /// <summary>
        /// Runs code inspections 
        /// </summary>
        /// <param name="parameter"></param>
        public override async void Execute(object parameter)
        {
            _presenter.Show();
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