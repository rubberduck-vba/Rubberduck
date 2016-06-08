using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.FindSymbol;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that lets user search for and navigate to any identifier.
    /// </summary>
    [ComVisible(false)]
    public class FindSymbolCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly DeclarationIconCache _iconCache;
        private readonly NavigateCommand _navigateCommand = new NavigateCommand();

        public FindSymbolCommand(VBE vbe, RubberduckParserState state, DeclarationIconCache iconCache)
        {
            _vbe = vbe;
            _state = state;
            _iconCache = iconCache;
        }

        public override void Execute(object parameter)
        {
            var viewModel = new FindSymbolViewModel(_state.AllDeclarations.Where(item => !item.IsBuiltIn), _iconCache);
            using (var view = new FindSymbolDialog(viewModel))
            {
                viewModel.Navigate += (sender, e) => { _navigateCommand.Execute(e); view.Hide(); };
                view.ShowDialog();
            }
        }
    }
}
