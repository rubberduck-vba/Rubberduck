using System.Linq;
using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Common;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI.FindSymbol;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that lets user search for and navigate to any identifier.
    /// </summary>
    [ComVisible(false)]
    public class FindSymbolCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly DeclarationIconCache _iconCache;
        private readonly NavigateCommand _navigateCommand = new NavigateCommand();

        public FindSymbolCommand(IVBE vbe, RubberduckParserState state, DeclarationIconCache iconCache) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _state = state;
            _iconCache = iconCache;
        }

        public override RubberduckHotkey Hotkey
        {
            get { return RubberduckHotkey.FindSymbol; }
        }

        protected override void ExecuteImpl(object parameter)
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
