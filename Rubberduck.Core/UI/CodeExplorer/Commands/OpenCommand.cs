using System;
using System.Collections.Generic;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Navigation.CodeExplorer;

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

        public OpenCommand(INavigateCommand openCommand)
        {
            _openCommand = openCommand;
        }

        public sealed override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        protected override bool EvaluateCanExecute(object parameter)
        {
            return base.EvaluateCanExecute(parameter) && 
                   ((CodeExplorerItemViewModel)parameter).QualifiedSelection.HasValue;
        }

        protected override void OnExecute(object parameter)
        {
            if (!EvaluateCanExecute(parameter))
            {
                return;
            }

            // ReSharper disable once PossibleInvalidOperationException - tested above.
            _openCommand.Execute(((CodeExplorerItemViewModel)parameter).QualifiedSelection.Value.GetNavitationArgs());
        }
    }
}
