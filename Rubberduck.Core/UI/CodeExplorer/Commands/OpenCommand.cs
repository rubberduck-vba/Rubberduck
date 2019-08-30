using System;
using System.Collections.Generic;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Navigation.CodeExplorer;
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

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public sealed override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return ((CodeExplorerItemViewModel)parameter).QualifiedSelection.HasValue;
        }

        protected override void OnExecute(object parameter)
        {
            if (!CanExecute(parameter))
            {
                return;
            }

            // ReSharper disable once PossibleInvalidOperationException - tested above.
            _openCommand.Execute(((CodeExplorerItemViewModel)parameter).QualifiedSelection.Value.GetNavitationArgs());
        }
    }
}
