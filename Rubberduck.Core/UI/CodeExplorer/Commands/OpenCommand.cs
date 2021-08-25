using System;
using System.Collections.Generic;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class OpenCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerProjectViewModel),
            typeof(CodeExplorerCustomFolderViewModel),
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel)
        };

        private readonly INavigateCommand _openCommand;

        public OpenCommand(
            INavigateCommand openCommand, 
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _openCommand = openCommand;

            AddToCanExecuteEvaluation(EvaluateCanExecute);
        }

        public sealed override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        private bool EvaluateCanExecute(object parameter)
        {
            return (parameter as CodeExplorerItemViewModel)?.QualifiedSelection.HasValue ?? false;
        }

        protected override void OnExecute(object parameter)
        {
            if (parameter is CodeExplorerItemViewModel item && item.QualifiedSelection.HasValue)
            {
                _openCommand.Execute(item.QualifiedSelection.Value.GetNavitationArgs());
            }
        }
    }
}
