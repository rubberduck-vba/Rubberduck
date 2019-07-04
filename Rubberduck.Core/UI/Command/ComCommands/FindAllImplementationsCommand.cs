using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.ComCommands
{
    /// <summary>
    /// A command that finds all implementations of a specified method, or of the active interface module.
    /// </summary>
    [ComVisible(false)]
    public class FindAllImplementationsCommand : ComCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IVBE _vbe;
        private readonly FindAllImplementationsService _finder;

        public FindAllImplementationsCommand(
            RubberduckParserState state, IVBE vbe, ISearchResultsWindowViewModel viewModel, 
            FindAllImplementationsService finder, IVbeEvents vbeEvents)
            : base(vbeEvents)
        {
            _finder = finder;
            _state = state;
            _vbe = vbe;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready)
            {
                return false;
            }

            using (var codePane = _vbe.ActiveCodePane)
            {
                if (codePane == null || codePane.IsWrappingNullReference)
                {
                    return false;
                }

                var target = FindTarget(parameter);
                return _finder.CanFind(target);
            }
        }

        protected override void OnExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready)
            {
                return;
            }

            var declaration = FindTarget(parameter);
            if (declaration == null)
            {
                return;
            }

            _finder.FindAllImplementations(declaration);
        }

        private Declaration FindTarget(object parameter)
        {
            if (parameter is Declaration declaration)
            {
                return declaration;
            }

            using (var activePane = _vbe.ActiveCodePane)
            {
                return _state.FindSelectedDeclaration(activePane);
            }
        }
    }
}
