using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public abstract class CodeExplorerCommandBase : CommandBase
    {
        protected CodeExplorerCommandBase()
        {
            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public abstract IEnumerable<Type> ApplicableNodeTypes { get; }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return parameter != null && ApplicableNodeTypes.Contains(parameter.GetType());
        }
    }
}
