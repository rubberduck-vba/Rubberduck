using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
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
        private readonly IRubberduckParser _parser;
        private readonly SearchResultIconCache _iconCache;

        public FindSymbolCommand(VBE vbe, IRubberduckParser parser, SearchResultIconCache iconCache)
        {
            _vbe = vbe;
            _parser = parser;
            _iconCache = iconCache;
        }

        public override void Execute(object parameter)
        {
            var progress = new ParsingProgressPresenter();
            var result = progress.Parse(_parser, _vbe.ActiveVBProject);
            var declarations = result.Declarations;
            var vm = new FindSymbolViewModel(declarations.Items.Where(item => !item.IsBuiltIn), _iconCache);
            using (var view = new FindSymbolDialog(vm))
            {
                view.ShowDialog();
            }
        }
    }
}