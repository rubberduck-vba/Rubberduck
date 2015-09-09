using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
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
        private readonly VBE _vbe;
        private readonly AddIn _addin;
        private readonly IInspector _inspector;
        private readonly IRubberduckParser _parser;

        public RunCodeInspectionsCommand(VBE vbe, AddIn addin, IInspector inspector, IRubberduckParser parser)
        {
            _vbe = vbe;
            _addin = addin;
            _inspector = inspector;
            _parser = parser;
        }

        /// <summary>
        /// Runs code inspections 
        /// </summary>
        /// <param name="parameter"></param>
        public override async void Execute(object parameter)
        {
            var factory = new CodePaneWrapperFactory();
            var viewModel = new InspectionResultsViewModel(_inspector, factory, _vbe);
            var presenter = new CodeInspectionsDockablePresenter(_inspector, _vbe, _addin, new CodeInspectionsWindow(viewModel), factory);
            presenter.Show();
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