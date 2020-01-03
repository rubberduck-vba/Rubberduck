using System.Runtime.InteropServices;
using Rubberduck.Common;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.FindSymbol;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.ComCommands
{
    /// <summary>
    /// A command that lets user search for and navigate to any identifier.
    /// </summary>
    [ComVisible(false)]
    public class FindSymbolCommand : ComCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly DeclarationIconCache _iconCache;
        private readonly NavigateCommand _navigateCommand;

        public FindSymbolCommand(
            RubberduckParserState state, 
            ISelectionService selectionService,
            DeclarationIconCache iconCache, 
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _state = state;
            _iconCache = iconCache;

            _navigateCommand = new NavigateCommand(selectionService);
        }

        protected override void OnExecute(object parameter)
        {
            var viewModel = new FindSymbolViewModel(_state.AllUserDeclarations, _iconCache);
            var view = new FindSymbolDialog(viewModel);
            {
                viewModel.Navigate += (sender, e) => { view.Hide(); };
                viewModel.Navigate += OnDialogNavigate;
                view.ShowDialog();
                _navigateCommand.Execute(_selected);
            }
        }

        private NavigateCodeEventArgs _selected;
        private void OnDialogNavigate(object sender, NavigateCodeEventArgs e)
        {
            _selected = e;
        }
    }
}
