using System.Runtime.InteropServices;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.CodeExplorer;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.Command.ComCommands
{
    /// <summary>
    /// A command that displays a popup near the cursor, owned by the Code Explorer WPF UserControl.
    /// </summary>
    [ComVisible(false)]
    public class PeekDefinitionCommand : ComCommandBase
    {
        private readonly IPeekDefinitionPopupProvider _provider;
        private readonly ISelectedDeclarationProvider _selection;

        public PeekDefinitionCommand(CodeExplorerDockablePresenter codeExplorer, IVbeEvents vbeEvents, ISelectedDeclarationProvider selection)
            : base(vbeEvents)
        {
            _provider = (codeExplorer.UserControl as CodeExplorerWindow)?.ViewModel;
            _selection = selection;
            AddToCanExecuteEvaluation(CanExecuteInternal);
        }

        private bool CanExecuteInternal(object parameter)
        {
            if (parameter is ModuleDeclaration || parameter is ModuleBodyElementDeclaration || parameter is VariableDeclaration || parameter is ValuedDeclaration)
            {
                return true;
            }

            return _selection.SelectedDeclaration() != null;
        }

        protected override void OnExecute(object parameter)
        {
            if (parameter is Declaration target)
            {
                _provider.PeekDefinition(target);
            }
            else
            {
                var selection = _selection.SelectedDeclaration();
                if (selection != null)
                {
                    _provider.PeekDefinition(selection);
                }
            }
        }
    }
}