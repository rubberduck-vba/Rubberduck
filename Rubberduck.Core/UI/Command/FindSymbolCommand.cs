using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Common;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Parsing.VBA;
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
        private readonly NavigateCommand _navigateCommand;

        public FindSymbolCommand(IVBE vbe, RubberduckParserState state, DeclarationIconCache iconCache) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _state = state;
            _iconCache = iconCache;

            _navigateCommand = new NavigateCommand(_state.ProjectsProvider);
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
