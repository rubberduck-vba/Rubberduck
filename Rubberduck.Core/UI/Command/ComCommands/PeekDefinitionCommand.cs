using System.Runtime.InteropServices;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.CodeExplorer;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.ComCommands
{
    /// <summary>
    /// A command that displays a popup near the cursor, owned by the Code Explorer WPF UserControl.
    /// </summary>
    [ComVisible(false)]
    public class PeekDefinitionCommand : ComCommandBase
    {
        public PeekDefinitionCommand(CodeExplorerDockablePresenter codeExplorer, IVbeEvents vbeEvents, ISelectedDeclarationProvider selection)
            : base(vbeEvents)
        {
            PopupProvider = (codeExplorer.UserControl as CodeExplorerWindow)?.ViewModel;
            SelectedDeclarationProvider = selection;
            AddToCanExecuteEvaluation(EvaluateCanExecute);
        }

        protected ISelectedDeclarationProvider SelectedDeclarationProvider { get; }
        protected IPeekDefinitionPopupProvider PopupProvider { get; }

        private bool EvaluateCanExecute(object parameter)
        {
            if (parameter is ModuleDeclaration || parameter is ModuleBodyElementDeclaration || parameter is VariableDeclaration || parameter is ValuedDeclaration)
            {
                return true;
            }

            return SelectedDeclarationProvider.SelectedDeclaration() != null;
        }

        protected override void OnExecute(object parameter)
        {
            if (parameter is Declaration target)
            {
                PopupProvider.PeekDefinition(target);
            }
            else
            {
                var selection = SelectedDeclarationProvider.SelectedDeclaration();
                if (selection != null)
                {
                    PopupProvider.PeekDefinition(selection);
                }
            }
        }
    }

    public class ProjectExplorerPeekDefinitionCommand : PeekDefinitionCommand
    {
        public ProjectExplorerPeekDefinitionCommand(CodeExplorerDockablePresenter codeExplorer, IVbeEvents vbeEvents, ISelectedDeclarationProvider selection)
            : base(codeExplorer, vbeEvents, selection)
        {}

        protected override void OnExecute(object parameter)
        {
            var module = SelectedDeclarationProvider.SelectedProjectExplorerModule();
            base.OnExecute(module);
        }
    }
}