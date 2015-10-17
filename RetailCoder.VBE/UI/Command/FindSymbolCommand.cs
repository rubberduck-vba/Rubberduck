using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.UI.FindSymbol;
using Rubberduck.UI.ParserProgress;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that lets user search for and navigate to any identifier.
    /// </summary>
    [ComVisible(false)]
    public class FindSymbolCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly IParsingProgressPresenter _parserProgress;
        private readonly DeclarationIconCache _iconCache;
        private readonly NavigateCommand _navigateCommand = new NavigateCommand();

        public FindSymbolCommand(VBE vbe, IParsingProgressPresenter parserProgress, DeclarationIconCache iconCache)
        {
            _vbe = vbe;
            _parserProgress = parserProgress;
            _iconCache = iconCache;
        }

        public override void Execute(object parameter)
        {
            var result = _parserProgress.Parse(_vbe.ActiveVBProject);
            var declarations = result.Declarations;
            var viewModel = new FindSymbolViewModel(declarations.Items.Where(item => !item.IsBuiltIn), _iconCache);
            using (var view = new FindSymbolDialog(viewModel))
            {
                viewModel.Navigate += (sender, e) => { _navigateCommand.Execute(e); view.Hide(); };
                view.ShowDialog();
            }
        }
    }
}